// xlsx-grow.js — przy zapisie: zakresy w pliku rosną razem z danymi dopisanymi pod spodem
// i z prawej (Etap 1a „nowe wiersze”, 1b „nowe kolumny”, 2026-11-10).
//
// Dlaczego: xlsx-patch.js dopisuje nowe wiersze do <sheetData>, ale zakresy opisane gdzie
// indziej zostawały przy starym końcu danych. W Excelu dopisany wiersz wypadał wtedy poza
// tabelę (bez jej stylu, filtra i formuł kolumn), poza autofiltr, formatowanie warunkowe
// i listy wyboru. Excel przy wpisywaniu tuż pod danymi rozszerza je sam — tu robimy to samo.
//
// Zasada (jak Excel przy dopisywaniu): rośnie zakres, który KOŃCZY SIĘ na ostatnim wierszu
// danych z otwartego pliku (fromRow) i ma co najmniej dwa wiersze — tak samo w bok: kończy się
// na ostatniej kolumnie danych (fromCol) i ma co najmniej dwie kolumny. Nie ruszamy:
//   • zakresów sięgających dalej niż dane (lista wyboru do wiersza 200 — już obejmuje nowe),
//   • pojedynczych komórek w ostatnim wierszu (np. jedna komórka sumy),
//   • całych kolumn (A:A) i zakresów kończących się wyżej,
//   • tabel z wierszem sumy w dół (Excel wstawia wtedy NAD sumą — to wstawianie w środku, Etap 4).
// Nowa kolumna tabeli Excela = nowy <tableColumn> z nazwą równą nagłówkowi w arkuszu (Excel
// wymaga, by nazwy się zgadzały — inaczej „naprawia” plik).
// Tylko tekst poza <sheetData> (jak xlsx-patch.js: bez DOMParsera na całym arkuszu).
// Wiersze i kolumny 1-based (A = 1), jak w adresach Excela.
// g = { fromRow, toRow, fromCol, toCol, tableCols: { nazwaTabeli(lower): [nazwy nowych kolumn] } }

const swbGrow = (() => {
  const CELL_RE = /^(\$?)([A-Z]{1,3})(\$?)(\d+)$/;
  const colNum = (letters) => { let n = 0; for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64); return n; };
  const colLetters = (n) => { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  // wywołania z dwoma liczbami (fromRow, toRow) — jak w Etapie 1a
  const asG = (a, b) => (a && typeof a === "object" ? a : { fromRow: a, toRow: b });
  const rowsOn = (g) => g.toRow > g.fromRow && g.fromRow >= 1;
  const colsOn = (g) => g.toCol > g.fromCol && g.fromCol >= 1;

  // „A3:H42” → rośnie w dół (koniec na fromRow, ≥ 2 wiersze) i w bok (koniec na fromCol, ≥ 2 kolumny).
  function growRange(range, a, b) {
    const g = asG(a, b);
    const parts = String(range).split(":");
    if (parts.length !== 2) return range; // pojedyncza komórka albo coś nietypowego
    const m1 = CELL_RE.exec(parts[0]);
    const m2 = CELL_RE.exec(parts[1]);
    if (!m1 || !m2) return range; // A:A, 3:3, odwołania do innych arkuszy…
    let r2 = Number(m2[4]);
    let c2 = colNum(m2[2]);
    const r1 = Number(m1[4]);
    const c1 = colNum(m1[2]);
    let changed = false;
    if (rowsOn(g) && r2 === g.fromRow && r1 < r2) { r2 = g.toRow; changed = true; }
    if (colsOn(g) && c2 === g.fromCol && c1 < c2) { c2 = g.toCol; changed = true; }
    if (!changed) return range;
    return `${parts[0]}:${m2[1]}${colLetters(c2)}${m2[3]}${r2}`;
  }

  // sqref = lista zakresów rozdzielonych spacjami
  function growSqref(sqref, a, b) {
    const g = asG(a, b);
    return String(sqref).split(/\s+/).filter(Boolean).map((r) => growRange(r, g)).join(" ");
  }

  function growAttr(tagXml, attr, g, fn = growRange) {
    return tagXml.replace(new RegExp(`(\\s${attr}=")([^"]*)(")`), (m, p1, v, p3) => p1 + fn(v, g) + p3);
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

  // <dimension ref="A1:K42"/> — koniec co najmniej do toRow / toCol (Excel liczy z niego „używany zakres”)
  function growDimension(before, g) {
    return before.replace(/(<(?:\w+:)?dimension\s[^>]*\bref=")([^"]*)(")/, (m, p1, ref, p3) => {
      const parts = ref.split(":");
      const end = CELL_RE.exec(parts[parts.length - 1]);
      if (!end) return m;
      const r = Math.max(Number(end[4]), rowsOn(g) ? g.toRow : 0, g.dimRow || 0);
      const c = Math.max(colNum(end[2]), colsOn(g) ? g.toCol : 0);
      if (r === Number(end[4]) && c === colNum(end[2])) return m;
      return `${p1}${parts[0]}:${colLetters(c)}${r}${p3}`;
    });
  }

  function growSheetXml(xml, a, b) {
    const g = asG(a, b);
    if (!rowsOn(g) && !colsOn(g) && !g.dimRow) return xml;
    const parts = splitOutsideSheetData(xml);
    if (!parts) return xml;
    const before = growDimension(parts.before, g);
    let after = parts.after;
    if (!rowsOn(g) && !colsOn(g)) return before + parts.data + after; // tylko wiersz za odstępem: sam wymiar
    // autofiltr arkusza
    after = after.replace(/<(?:\w+:)?autoFilter\s[^>]*>/g, (tag) => growAttr(tag, "ref", g));
    // formatowanie warunkowe i listy wyboru (zwykłe: atrybut sqref)
    after = after.replace(/<(?:\w+:)?conditionalFormatting\s[^>]*>/g, (tag) => growAttr(tag, "sqref", g, growSqref));
    after = after.replace(/<(?:\w+:)?dataValidation\s[^>]*>/g, (tag) => growAttr(tag, "sqref", g, growSqref));
    // rozszerzenia x14 (nowsze typy reguł): zakres w <xm:sqref>…</xm:sqref>
    after = after.replace(/(<(\w+:)?sqref>)([^<]*)(<\/\2?sqref>)/g, (m, open, _p, v, close) => open + growSqref(v, g) + close);
    return before + parts.data + after;
  }

  const escAttr = (s) => String(s).replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\n/g, "_x000a_");

  // Tabela Excela (xl/tables/tableN.xml): ref i autoFilter rosną, gdy tabela kończy się na
  // ostatnim wierszu / ostatniej kolumnie danych. W dół — bez wiersza sumy (totalsRowCount ≥ 1).
  // W bok — nowe <tableColumn> z nazwami z g.tableCols[nazwa tabeli]; bez nazw tabela nie rośnie
  // w bok (lepiej poza tabelą niż tabela z kolumną bez nazwy, którą Excel by „naprawiał”).
  function growTableXml(xml, a, b) {
    const g0 = asG(a, b);
    const root = /<(?:\w+:)?table\s[^>]*>/.exec(xml);
    if (!root) return xml;
    const tag = root[0];
    const totals = /\stotalsRowCount="(\d+)"/.exec(tag);
    const name = ((/\sname="([^"]*)"/.exec(tag) || [])[1] || "").toLowerCase();
    const disp = ((/\sdisplayName="([^"]*)"/.exec(tag) || [])[1] || "").toLowerCase();
    const names = g0.tableCols ? (g0.tableCols[disp] || g0.tableCols[name] || null) : null;
    const g = {
      fromRow: totals && Number(totals[1]) > 0 ? 0 : g0.fromRow,
      toRow: g0.toRow,
      fromCol: names && names.length && g0.toCol - g0.fromCol === names.length ? g0.fromCol : 0,
      toCol: g0.toCol,
    };
    if (!rowsOn(g) && !colsOn(g)) return xml;
    const ref = /\sref="([^"]*)"/.exec(tag);
    if (!ref) return xml;
    const grownRef = growRange(ref[1], g);
    if (grownRef === ref[1]) return xml;
    const colsGrew = colsOn(g) && CELL_RE.exec(grownRef.split(":")[1])[2] !== CELL_RE.exec(ref[1].split(":")[1])[2];
    let out = xml.slice(0, root.index) + growAttr(tag, "ref", g) + xml.slice(root.index + tag.length);
    out = out.replace(/<(?:\w+:)?autoFilter\s[^>]*>/, (t) => growAttr(t, "ref", g));
    if (colsGrew) {
      const tc = /<((?:\w+:)?)tableColumns\b([^>]*)>([\s\S]*?)<\/\1tableColumns>/.exec(out);
      if (!tc) return xml; // nietypowa tabela — bez zmian w całości
      const p = tc[1];
      const ids = [...tc[3].matchAll(/<(?:\w+:)?tableColumn\b[^>]*\sid="(\d+)"/g)].map((m) => Number(m[1]));
      let next = (ids.length ? Math.max(...ids) : 0) + 1;
      const added = names.map((n) => `<${p}tableColumn id="${next++}" name="${escAttr(n)}"/>`).join("");
      const count = ids.length + names.length;
      const attrs = /\scount="\d+"/.test(tc[2]) ? tc[2].replace(/\scount="\d+"/, ` count="${count}"`) : `${tc[2]} count="${count}"`;
      out = out.slice(0, tc.index) + `<${p}tableColumns${attrs}>${tc[3]}${added}</${p}tableColumns>` + out.slice(tc.index + tc[0].length);
    }
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

  // Cały plik: dla każdego arkusza z g — arkusz i jego tabele.
  async function growZip(zip, sheetPaths, growBySheet) {
    for (const sheetName of Object.keys(growBySheet || {})) {
      const g = growBySheet[sheetName];
      const path = sheetPaths[sheetName];
      if (!g || !path || !(rowsOn(g) || colsOn(g)) || !zip.file(path)) continue;
      const xml = await zip.file(path).async("string");
      const grown = growSheetXml(xml, g);
      if (grown !== xml) zipWriteUtf8Xml(zip, path, grown);
      const relsPath = path.replace(/([^/]+)$/, "_rels/$1.rels");
      const rels = zip.file(relsPath) ? await zip.file(relsPath).async("string") : "";
      for (const tpath of tablePathsFromRels(rels, path)) {
        if (!zip.file(tpath)) continue;
        const txml = await zip.file(tpath).async("string");
        const tgrown = growTableXml(txml, g);
        if (tgrown !== txml) zipWriteUtf8Xml(zip, tpath, tgrown);
      }
    }
  }

  return { growRange, growSqref, growSheetXml, growTableXml, tablePathsFromRels, growZip };
})();

if (typeof module !== "undefined" && module.exports) module.exports = swbGrow;
