// @ts-check
// xlsx-layout.js — szerokości kolumn i wysokości wierszy ustawione w aplikacji → plik (docs/OD-ZERA.md, poz. 1).
//
// layout = { cols: { indeksKolumny0: szerokośćWZnakach }, rows: { indeksWiersza0: wysokośćWPunktach } }
// Tekstowo, jak xlsx-patch.js: <cols> poza <sheetData> i atrybuty <row> — reszta arkusza bajt w bajt.
// Zakres <col min max> obejmujący zmienianą kolumnę dzielimy na części z tymi samymi atrybutami
// (styl kolumny, ukrycie zostają), a zmieniona dostaje width + customWidth="1" — jak zapisuje Excel.
// Przeliczenie (Excel Mac i SheetJS, sprawdzone 2026-11-15): szerokość na ekranie = znaki × 6 + 5 pkt,
// w pliku width = px / 6 (np. 125 px → 20.83203125); wysokość wiersza: px = pkt (jak SheetJS hpx).

const swbLayout = (() => {
  const colLetters = (n) => { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  // px przy 100% → width z pliku (1/256 znaku, jak Excel) i z powrotem
  const pxToWidth = (px) => Math.floor((Math.max(0, px) / 6) * 256) / 256;
  const widthToPx = (w) => Math.round(w * 6);
  const pxToPt = (px) => Math.round(Math.max(0, px) * 4) / 4;
  const num = (v) => String(Number(v));
  const attr = (tag, name) => { const m = String(tag).match(new RegExp(`\\s${name}="([^"]*)"`)); return m ? m[1] : null; };
  const setAttr = (tag, name, value) => (attr(tag, name) == null
    ? tag.replace(/\s*(\/?)>$/, ` ${name}="${value}"$1>`)
    : tag.replace(new RegExp(`(\\s${name}=")[^"]*(")`), `$1${value}$2`));

  function patchCols(xml, cols) {
    const keys = Object.keys(cols || {}).map(Number).filter((c) => Number.isInteger(c) && c >= 0 && Number.isFinite(cols[c]));
    if (!keys.length) return xml;
    const m = xml.match(/<((?:[\w-]+:)?)cols>([\s\S]*?)<\/\1cols>|<((?:[\w-]+:)?)cols\s*\/>/);
    const pfx = m ? (m[1] ?? m[3] ?? "") : ((xml.match(/<((?:[\w-]+:)?)sheetData\b/) || [])[1] || "");
    let spans = m && m[2] ? [...m[2].matchAll(/<(?:[\w-]+:)?col\b[^>]*\/?>(?:<\/(?:[\w-]+:)?col>)?/g)].map((x) => {
      const raw = x[0].replace(/><\/(?:[\w-]+:)?col>$/, "/>");
      return { min: Number(attr(raw, "min")), max: Number(attr(raw, "max")), raw };
    }) : [];
    keys.forEach((c0) => {
      const n = c0 + 1;
      const width = num(cols[c0]);
      const i = spans.findIndex((s) => n >= s.min && n <= s.max);
      if (i < 0) {
        spans.push({ min: n, max: n, raw: `<${pfx}col min="${n}" max="${n}" width="${width}" customWidth="1"/>` });
        return;
      }
      const s = spans[i];
      const part = (a, b) => ({ min: a, max: b, raw: setAttr(setAttr(s.raw, "min", String(a)), "max", String(b)) });
      const mid = part(n, n);
      mid.raw = setAttr(setAttr(mid.raw, "width", width), "customWidth", "1");
      const pieces = [];
      if (s.min < n) pieces.push(part(s.min, n - 1));
      pieces.push(mid);
      if (n < s.max) pieces.push(part(n + 1, s.max));
      spans.splice(i, 1, ...pieces);
    });
    spans.sort((a, b) => a.min - b.min);
    const block = `<${pfx}cols>${spans.map((s) => s.raw).join("")}</${pfx}cols>`;
    if (m) return xml.replace(m[0], block);
    // schemat: <cols> po sheetFormatPr (albo sheetViews / sheetPr / dimension), przed <sheetData>
    const sd = xml.search(/<(?:[\w-]+:)?sheetData\b/);
    return sd < 0 ? xml : xml.slice(0, sd) + block + xml.slice(sd);
  }

  function patchRows(xml, rows) {
    const keys = Object.keys(rows || {}).map(Number).filter((r) => Number.isInteger(r) && r >= 0 && Number.isFinite(rows[r])).sort((a, b) => a - b);
    if (!keys.length) return xml;
    // pusty arkusz: <sheetData/> → para znaczników, żeby było gdzie wstawić wiersz
    xml = xml.replace(/<((?:[\w-]+:)?)sheetData(\s[^>]*)?\/>/, "<$1sheetData$2></$1sheetData>");
    const open = xml.match(/<((?:[\w-]+:)?)sheetData\b[^>]*>/);
    if (!open) return xml;
    const pfx = open[1];
    const start = open.index + open[0].length;
    const end = xml.indexOf(`</${pfx}sheetData>`, start);
    if (end < 0) return xml;
    let inner = xml.slice(start, end);
    keys.forEach((r0) => {
      const r = r0 + 1;
      const ht = num(rows[r0]);
      const re = new RegExp(`<${pfx}row\\b[^>]*\\sr="${r}"[^>]*?(/?)>`);
      const hit = inner.match(re);
      if (hit) {
        inner = inner.replace(hit[0], setAttr(setAttr(hit[0], "ht", ht), "customHeight", "1"));
        return;
      }
      const tag = `<${pfx}row r="${r}" ht="${ht}" customHeight="1"/>`;
      const later = [...inner.matchAll(new RegExp(`<${pfx}row\\b[^>]*\\sr="(\\d+)"`, "g"))].find((x) => Number(x[1]) > r);
      inner = later ? inner.slice(0, later.index) + tag + inner.slice(later.index) : inner + tag;
    });
    return xml.slice(0, start) + inner + xml.slice(end);
  }

  function patchSheetLayout(xml, layout) {
    if (!layout) return xml;
    return patchRows(patchCols(xml, layout.cols), layout.rows);
  }

  // Do zapisu: layouts = { arkusz: layout }, ścieżki arkuszy z resolveSheetPaths
  async function applyToZip(zip, sheetPaths, layouts, write) {
    for (const name of Object.keys(layouts || {})) {
      const l = layouts[name];
      const path = sheetPaths[name];
      if (!l || !path || !zip.file(path)) continue;
      if (!Object.keys(l.cols || {}).length && !Object.keys(l.rows || {}).length) continue;
      const xml = await zip.file(path).async("string");
      const next = patchSheetLayout(xml, l);
      if (next !== xml) write(zip, path, next);
    }
  }

  return { pxToWidth, widthToPx, pxToPt, patchCols, patchRows, patchSheetLayout, applyToZip, colLetters };
})();

if (typeof module !== "undefined" && module.exports) module.exports = swbLayout;
