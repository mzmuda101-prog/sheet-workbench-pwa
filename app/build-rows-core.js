// [EN] Pure buildRows logic for Web Worker — mirrors table.js buildRows; main thread keeps sync fallback.
(function (global) {
  "use strict";

  const INDEXED_COLORS = [
    "#000000","#FFFFFF","#FF0000","#00FF00","#0000FF","#FFFF00","#FF00FF","#00FFFF",
    "#000000","#FFFFFF","#FF0000","#00FF00","#0000FF","#FFFF00","#FF00FF","#00FFFF",
    "#800000","#008000","#000080","#808000","#800080","#008080","#C0C0C0","#808080",
    "#9999FF","#993366","#FFFFCC","#CCFFFF","#660066","#FF8080","#0066CC","#CCCCFF",
    "#000080","#FF00FF","#FFFF00","#00FFFF","#800080","#800000","#008080","#0000FF",
    "#00CCFF","#CCFFFF","#CCFFCC","#FFFF99","#99CCFF","#FF99CC","#CC99FF","#FFCC99",
    "#3366FF","#33CCCC","#99CC00","#FFCC00","#FF9900","#FF6600","#666699","#969696",
    "#003366","#339966","#003300","#333300","#993300","#993366","#333399","#333333",
    null, null,
  ];

  const THEME_COLORS = [
    "#FFFFFF", "#000000", "#E7E6E6", "#44546A",
    "#4472C4", "#ED7D31", "#A5A5A5", "#FFC000",
    "#5B9BD5", "#70AD47", "#0563C1", "#954F72",
  ];

  function normalizeHexColor(input) {
    if (!input) return null;
    const raw = String(input).replace(/^#/, "").trim();
    if (/^[A-Fa-f0-9]{8}$/.test(raw)) return `#${raw.slice(2)}`;
    if (/^[A-Fa-f0-9]{6}$/.test(raw)) return `#${raw}`;
    if (/^[A-Fa-f0-9]{3}$/.test(raw)) return `#${raw[0]}${raw[0]}${raw[1]}${raw[1]}${raw[2]}${raw[2]}`;
    return null;
  }

  function applyColorTint(hex, tint) {
    const m = /^#?([0-9a-fA-F]{6})$/.exec(hex || "");
    if (!m || !tint) return hex;
    const num = parseInt(m[1], 16);
    let r = ((num >> 16) & 255) / 255;
    let g = ((num >> 8) & 255) / 255;
    let b = (num & 255) / 255;
    const max = Math.max(r, g, b), min = Math.min(r, g, b);
    let h = 0, s = 0, l = (max + min) / 2;
    const d = max - min;
    if (d !== 0) {
      s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
      if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
      else if (max === g) h = (b - r) / d + 2;
      else h = (r - g) / d + 4;
      h /= 6;
    }
    l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
    l = Math.min(1, Math.max(0, l));
    let R, G, B;
    if (s === 0) { R = G = B = l; }
    else {
      const hue2rgb = (p, q, t) => {
        if (t < 0) t += 1;
        if (t > 1) t -= 1;
        if (t < 1 / 6) return p + (q - p) * 6 * t;
        if (t < 1 / 2) return q;
        if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
        return p;
      };
      const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
      const p = 2 * l - q;
      R = hue2rgb(p, q, h + 1 / 3); G = hue2rgb(p, q, h); B = hue2rgb(p, q, h - 1 / 3);
    }
    const toHex = (v) => Math.round(v * 255).toString(16).padStart(2, "0");
    return `#${toHex(R)}${toHex(G)}${toHex(B)}`;
  }

  function colorFromStyleNode(node, resolveTheme = false) {
    if (!node || typeof node !== "object") return null;
    const direct = normalizeHexColor(node.rgb ?? node.RGB);
    if (direct) return direct;
    if (resolveTheme) {
      const tint = Number(node.tint ?? node.Tint) || 0;
      const idx = node.indexed ?? node.Indexed;
      if (Number.isInteger(idx) && INDEXED_COLORS[idx]) {
        return tint ? applyColorTint(INDEXED_COLORS[idx], tint) : INDEXED_COLORS[idx];
      }
      const themeIdx = node.theme ?? node.Theme;
      if (Number.isInteger(themeIdx) && THEME_COLORS[themeIdx]) {
        return applyColorTint(THEME_COLORS[themeIdx], tint);
      }
    }
    return normalizeHexColor(node.auto) || null;
  }

  function isDefaultLikeFill(fill, fillColor) {
    if (!fill || typeof fill !== "object") return true;
    const patternType = String(fill.patternType || fill.PatternType || "none").toLowerCase();
    if (!patternType || patternType === "none") return true;
    if (!fillColor) return true;
    const normalized = String(fillColor).toUpperCase();
    if (normalized === "#FFFFFF" || normalized === "#FFFFFE") {
      const fg = fill.fgColor || fill.FgColor || null;
      const hasExplicitFgColor = !!(fg && typeof fg === "object" && (fg.rgb != null || fg.RGB != null || fg.theme != null || fg.Theme != null));
      return !(patternType === "solid" && hasExplicitFgColor);
    }
    if (normalized === "#000000") return true;
    return false;
  }

  function isDefaultLikeFontColor(fontColor) {
    if (!fontColor) return true;
    const normalized = String(fontColor).toUpperCase();
    return normalized === "#000000" || normalized === "#FFFFFF";
  }

  function isCustomAlignment(alignment) {
    if (!alignment || typeof alignment !== "object") return false;
    const horizontal = String(alignment.horizontal || alignment.Horizontal || "").toLowerCase();
    const vertical = String(alignment.vertical || alignment.Vertical || "").toLowerCase();
    const wrapText = !!(alignment.wrapText || alignment.wrap || alignment.WrapText);
    return !(!horizontal || horizontal === "general") || !(!vertical || vertical === "bottom") || wrapText;
  }

  function hasCustomBorder(border) {
    if (!border || typeof border !== "object") return false;
    return [border.top || border.Top, border.right || border.Right, border.bottom || border.Bottom, border.left || border.Left]
      .some((edge) => {
        const style = String(edge?.style || edge?.Style || "").toLowerCase();
        return !!style && style !== "none";
      });
  }

  function getWorkbookDefaultFontSize(wb) {
    const fonts = wb?.Styles?.Fonts || wb?.Styles?.fonts;
    const f0 = Array.isArray(fonts) ? fonts[0] : null;
    const sz = f0 ? Number(f0.sz ?? f0.Sz ?? f0.size ?? f0.Size) : 0;
    return sz > 0 ? sz : 11;
  }

  function excelFontToCssStack(name) {
    const n = String(name || "").trim();
    if (!n) return "";
    return `"${n}", sans-serif`;
  }

  function resolveXfStyleFromIndex(sIdx, wb) {
    if (!Number.isInteger(sIdx) || !wb || !wb.Styles) return null;
    const st = wb.Styles;
    const xfs = st.CellXf || st.CellXfs || st.cellXfs;
    const xf = Array.isArray(xfs) ? xfs[sIdx] : null;
    if (!xf) return null;
    const pick = (a, b) => {
      const v = a ?? b;
      const n = Number(v);
      return Number.isFinite(n) ? n : null;
    };
    const fontId = pick(xf.fontId, xf.fontid);
    const fillId = pick(xf.fillId, xf.fillid);
    const borderId = pick(xf.borderId, xf.borderid);
    const fonts = st.Fonts || st.fonts || [];
    const fills = st.Fills || st.fills || [];
    const borders = st.Borders || st.borders || [];
    return {
      fontId,
      font: fontId != null ? fonts[fontId] : null,
      fill: fillId != null ? fills[fillId] : null,
      border: borderId != null ? borders[borderId] : null,
      alignment: xf.alignment || xf.Alignment || null,
    };
  }

  function extractCellStyle(cell, wb, xfStyle = null) {
    const cellS = cell && cell.s && typeof cell.s === "object" ? cell.s : null;
    if (!cellS && !xfStyle) return null;
    const cellFillNode = cellS && (cellS.patternType != null || cellS.fgColor || cellS.FgColor)
      ? cellS
      : (cellS?.fill || cellS?.Fill || null);
    const fill = (xfStyle && xfStyle.fill) || cellFillNode || null;
    const border = (xfStyle && xfStyle.border) || cellS?.border || cellS?.Border || null;
    const alignment = (xfStyle && xfStyle.alignment) || cellS?.alignment || cellS?.Alignment || null;
    const rawFont = (xfStyle && xfStyle.font) || cellS?.font || cellS?.Font || null;
    const isDefaultFont = xfStyle ? xfStyle.fontId === 0 : false;
    const font = isDefaultFont ? null : rawFont;
    const fillColor = colorFromStyleNode(fill?.fgColor || fill?.FgColor || fill?.bgColor || fill?.BgColor, true);
    const fontColor = colorFromStyleNode(font?.color || font?.Color, true);
    const hasCustomFill = !isDefaultLikeFill(fill, fillColor);
    const fontColorIsDefaultLike = isDefaultLikeFontColor(fontColor);
    const hasCustomFontColor = !!fontColor && (!fontColorIsDefaultLike || hasCustomFill);
    const hasCustomAlign = isCustomAlignment(alignment);
    const hasBorder = hasCustomBorder(border);
    const fontName = (font && (font.name || font.Name || font.rFont)) || "";
    const rawFontSize = font ? Number(font.sz ?? font.Sz ?? font.size ?? font.Size) : 0;
    const baseFontSize = getWorkbookDefaultFontSize(wb);
    return {
      fillColor,
      hasCustomFill,
      fontColor,
      hasCustomFontColor,
      fontColorIsDefaultLike,
      fontFamily: fontName ? excelFontToCssStack(fontName) : "",
      fontScale: rawFontSize > 0 && baseFontSize > 0 ? rawFontSize / baseFontSize : 0,
      bold: !!(font && (font.bold || font.b || font.Bold)),
      italic: !!(font && (font.italic || font.i || font.Italic)),
      underline: !!(font && (font.underline || font.u || font.Underline)),
      horizontal: hasCustomAlign ? (alignment?.horizontal || alignment?.Horizontal || "") : "",
      vertical: hasCustomAlign ? (alignment?.vertical || alignment?.Vertical || "") : "",
      wrapText: hasCustomAlign ? !!(alignment && (alignment.wrapText || alignment.wrap || alignment.WrapText)) : false,
      hasBorder,
      border,
    };
  }

  function isMeaningfulSheetCell(cell) {
    if (!cell || typeof cell !== "object") return false;
    if (cell.f) return true;
    if (cell.l && (cell.l.Target || cell.l.target)) return true;
    if (Array.isArray(cell.c) && cell.c.length) return true;
    if (cell.v instanceof Date) return true;
    if (typeof cell.v === "number" && Number.isFinite(cell.v)) return true;
    if (typeof cell.v === "boolean") return true;
    if (typeof cell.v === "string" && cell.v.trim() !== "") return true;
    if (typeof cell.w === "string" && cell.w.trim() !== "") return true;
    return false;
  }

  function computeEffectiveSheetRange(sheet, headerRow) {
    const fallback = XLSX.utils.decode_range(sheet["!ref"]);
    const headerIndex0 = Math.max(0, (headerRow || 1) - 1);
    let minCol = fallback.e.c;
    let maxCol = fallback.s.c;
    let maxRow = headerIndex0;
    let found = false;
    Object.keys(sheet).forEach((key) => {
      if (!key || key[0] === "!") return;
      const cell = sheet[key];
      if (!isMeaningfulSheetCell(cell)) return;
      const ref = XLSX.utils.decode_cell(key);
      if (ref.r < headerIndex0) {
        minCol = Math.min(minCol, ref.c);
        maxCol = Math.max(maxCol, ref.c);
        found = true;
        return;
      }
      minCol = Math.min(minCol, ref.c);
      maxCol = Math.max(maxCol, ref.c);
      maxRow = Math.max(maxRow, ref.r);
      found = true;
    });
    for (let c = fallback.s.c; c <= fallback.e.c; c++) {
      const cell = sheet[XLSX.utils.encode_cell({ r: headerIndex0, c })];
      if (!isMeaningfulSheetCell(cell)) continue;
      minCol = Math.min(minCol, c);
      maxCol = Math.max(maxCol, c);
      found = true;
    }
    if (!found) return fallback;
    const merges = Array.isArray(sheet["!merges"]) ? sheet["!merges"] : [];
    merges.forEach((merge) => {
      if (!merge || !merge.s || !merge.e) return;
      const anchor = sheet[XLSX.utils.encode_cell({ r: merge.s.r, c: merge.s.c })];
      if (!isMeaningfulSheetCell(anchor)) return;
      minCol = Math.min(minCol, merge.s.c);
      maxCol = Math.max(maxCol, merge.e.c);
      if (merge.e.r >= headerIndex0) maxRow = Math.max(maxRow, merge.e.r);
    });
    return {
      s: { r: fallback.s.r, c: Math.max(fallback.s.c, minCol) },
      e: { r: Math.max(headerIndex0, maxRow), c: Math.max(Math.max(fallback.s.c, minCol), maxCol) },
    };
  }

  function makeHeadersUnique(headers, lang) {
    const label = lang === "en" ? "Column" : "Kolumna";
    const seen = new Map();
    return headers.map((header, index) => {
      const base = String(header || `${label} ${index + 1}`).trim() || `${label} ${index + 1}`;
      const count = seen.get(base) || 0;
      seen.set(base, count + 1);
      return count ? `${base} (${count + 1})` : base;
    });
  }

  function toDisplay(value) {
    if (value === null || value === undefined) return "";
    if (value instanceof Date) {
      const dd = String(value.getDate()).padStart(2, "0");
      const mm = String(value.getMonth() + 1).padStart(2, "0");
      const yy = String(value.getFullYear()).slice(-2);
      return `${dd}-${mm}-${yy}`;
    }
    return String(value);
  }

  function localizeDisplayedDate(value, shown, cell, locale) {
    if (!shown || typeof shown !== "string") return shown;
    if (!(value instanceof Date) && typeof value !== "number" && typeof value !== "string") return shown;
    const formatHint = String(cell?.z || cell?.w || "").toLowerCase();
    if (!/(mmmm|mmm|dddd|ddd)/i.test(formatHint) && !/\b(jan|feb|sty|lut|mar|monday|poniedz)/i.test(shown.toLowerCase())) {
      return shown;
    }
    let date = value instanceof Date ? value : null;
    if (!date && typeof value === "number" && Number.isFinite(value)) {
      date = new Date((value - 25569) * 86400000);
    }
    if (!date || Number.isNaN(date.getTime())) return shown;
    try {
      return new Intl.DateTimeFormat(locale || "pl-PL", { day: "2-digit", month: "short", year: "2-digit" })
        .format(date).replace(/\.$/g, "").replace(/\s+/g, " ").trim() || shown;
    } catch {
      return shown;
    }
  }

  function buildRowsCore(sheet, headerRow, wb, options) {
    const opts = options || {};
    const displayMode = opts.displayMode || "values";
    const lang = opts.lang || "pl";
    const locale = opts.locale || "pl-PL";
    const styleIndexMap = opts.styleIndexEntries ? new Map(opts.styleIndexEntries) : null;
    const originalRange = XLSX.utils.decode_range(sheet["!ref"]);
    const range = computeEffectiveSheetRange(sheet, headerRow);
    const colMeta = sheet["!cols"] || [];
    const rowMeta = sheet["!rows"] || [];
    const merges = Array.isArray(sheet["!merges"]) ? sheet["!merges"] : [];
    const xfStyleAt = (ref) => (styleIndexMap ? resolveXfStyleFromIndex(styleIndexMap.get(ref), wb) : null);
    const rawHeaders = [];
    const headerStyles = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const ref = XLSX.utils.encode_cell({ r: headerRow - 1, c });
      const cell = sheet[ref];
      const v = cell ? cell.v : null;
      rawHeaders.push(v ? String(v).trim() : XLSX.utils.encode_col(c));
      headerStyles.push(wb ? extractCellStyle(cell, wb, xfStyleAt(ref)) : null);
    }
    const headers = makeHeadersUnique(rawHeaders, lang);
    const duplicateHeaderCount = rawHeaders.length - new Set(rawHeaders).size;
    const rows = [];
    let formulaCount = 0;
    let formulaMissingResultCount = 0;
    let commentCount = 0;
    let hyperlinkCount = 0;
    for (let r = headerRow; r <= range.e.r; r++) {
      const values = [];
      const display = [];
      const cellStyles = [];
      let any = false;
      for (let c = range.s.c; c <= range.e.c; c++) {
        const ref = XLSX.utils.encode_cell({ r, c });
        const cell = sheet[ref];
        let v = cell ? cell.v : null;
        let shown = cell && cell.w ? String(cell.w) : toDisplay(v);
        shown = localizeDisplayedDate(v, shown, cell, locale);
        if (cell && cell.f) {
          formulaCount += 1;
          if (cell.v == null && cell.w == null) formulaMissingResultCount += 1;
        }
        if (cell && Array.isArray(cell.c) && cell.c.length) commentCount += 1;
        if (cell && cell.l && (cell.l.Target || cell.l.target)) hyperlinkCount += 1;
        if (displayMode === "formulas" && cell && cell.f) {
          v = "=" + cell.f;
          shown = v;
        }
        values.push(v);
        display.push(shown);
        cellStyles.push(wb ? extractCellStyle(cell, wb, xfStyleAt(ref)) : null);
        if (v !== null && v !== "") any = true;
      }
      if (!any) continue;
      rows.push({ values, display, rawValues: values, rowIndex0: r, cellStyles });
    }
    const colWidths = headers.map((_, i) => colMeta[range.s.c + i] || null);
    const rowHeights = {};
    for (let i = 0; i < rowMeta.length; i++) {
      if (rowMeta[i]) rowHeights[i] = rowMeta[i];
    }
    return {
      headers,
      headerStyles,
      rows,
      startCol: range.s.c,
      merges,
      colWidths,
      rowHeights,
      stats: {
        duplicateHeaderCount,
        formulaCount,
        formulaMissingResultCount,
        commentCount,
        hyperlinkCount,
        mergeRegions: merges.length,
        mergedCells: merges.reduce((sum, merge) => sum + ((merge.e.r - merge.s.r + 1) * (merge.e.c - merge.s.c + 1)), 0),
        hiddenColumns: colMeta.filter((meta) => meta && meta.hidden).length,
        hiddenRows: rowMeta.filter((meta) => meta && meta.hidden).length,
        sourceRange: XLSX.utils.encode_range(originalRange),
        effectiveRange: XLSX.utils.encode_range(range),
      },
    };
  }

  global.buildRowsCore = buildRowsCore;
})(typeof self !== "undefined" ? self : globalThis);
