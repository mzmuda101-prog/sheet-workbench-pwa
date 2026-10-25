// formula-engine.js — silnik formuł Excela (A1, zakresy, inne arkusze, tabele, nazwy).
//
// Czysty moduł bez DOM: parse → AST → evaluate / serialize / refsOf. Używa go aplikacja
// (wpisywanie formuł, przeliczanie komórek zależnych — formula-ui.js) i testy w Node,
// które liczą każdą formułę z prawdziwego pliku i porównują z wynikiem zapisanym przez Excela.
//
// Zasady jak w Excelu:
// - kolejność działań: zakres > minus jednoargumentowy > % > ^ > * / > + - > & > porównania
//   (=-2^2 daje 4, 2^3^2 daje 64 — lewostronnie),
// - puste = 0 w arytmetyce, "" w tekście; tekst-liczba („5”) jest liczbą w działaniu,
// - porównania: liczby < tekst < logiczne, tekst bez wielkości liter,
// - błędy (#DZIEL/0! itd.) przechodzą dalej, JEŻELI liczy tylko wybraną gałąź,
// - SUMA/ŚREDNIA z ZAKRESU pomijają tekst i wartości logiczne, z argumentu wprost — nie.
// Polskie nazwy funkcji (SUMA, JEŻELI, WYSZUKAJ.PIONOWO…), separator „;” i przecinek
// dziesiętny są przyjmowane na wejściu; do pliku idzie zawsze zapis Excela (angielski, „,”,
// prefiks _xlfn. dla funkcji nowszych niż Excel 2007).

(function (root) {
  "use strict";

  // ── błędy ──────────────────────────────────────────────────────────────────
  const ERR = {};
  ["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#SPILL!", "#CALC!"].forEach((code) => {
    ERR[code] = Object.freeze({ err: code });
  });
  const isErr = (v) => !!(v && typeof v === "object" && typeof v.err === "string");
  const errOf = (code) => ERR[code] || Object.freeze({ err: code });
  // Polskie nazwy błędów (wyświetlanie po polsku i wpisywanie)
  const ERR_PL = { "#NULL!": "#ZERO!", "#DIV/0!": "#DZIEL/0!", "#VALUE!": "#ARG!", "#REF!": "#ADR!", "#NAME?": "#NAZWA?", "#NUM!": "#LICZBA!", "#N/A": "#N/D", "#SPILL!": "#ROZLANIE!", "#CALC!": "#OBL!" };
  const ERR_FROM_PL = Object.fromEntries(Object.entries(ERR_PL).map(([en, pl]) => [pl, en]));
  // Wartość nieznana (funkcja spoza listy, której Excel jednak może znać) — komórka zachowuje
  // ostatni wynik z pliku, a Excel przeliczy ją przy otwarciu.
  const UNSUPPORTED = Object.freeze({ unsupported: true });

  // ── nazwy funkcji: angielskie (plik) ↔ polskie (Excel PL) ───────────────────
  const PL_NAMES = {
    SUM: "SUMA", AVERAGE: "ŚREDNIA", COUNT: "ILE.LICZB", COUNTA: "ILE.NIEPUSTYCH", COUNTBLANK: "LICZ.PUSTE",
    COUNTIF: "LICZ.JEŻELI", COUNTIFS: "LICZ.WARUNKI", SUMIF: "SUMA.JEŻELI", SUMIFS: "SUMA.WARUNKÓW",
    AVERAGEIF: "ŚREDNIA.JEŻELI", AVERAGEIFS: "ŚREDNIA.WARUNKÓW", MIN: "MIN", MAX: "MAKS", MINIFS: "MIN.WARUNKÓW",
    MAXIFS: "MAKS.WARUNKÓW", IF: "JEŻELI", IFS: "WARUNKI", IFERROR: "JEŻELI.BŁĄD", IFNA: "JEŻELI.ND", AND: "ORAZ",
    OR: "LUB", NOT: "NIE", XOR: "XOR", TRUE: "PRAWDA", FALSE: "FAŁSZ", SWITCH: "PRZEŁĄCZ", CHOOSE: "WYBIERZ",
    ROUND: "ZAOKR", ROUNDUP: "ZAOKR.GÓRA", ROUNDDOWN: "ZAOKR.DÓŁ", INT: "ZAOKR.DO.CAŁK", TRUNC: "LICZBA.CAŁK",
    ABS: "MODUŁ.LICZBY", MOD: "MOD", POWER: "POTĘGA", SQRT: "PIERWIASTEK", PRODUCT: "ILOCZYN",
    SUMPRODUCT: "SUMA.ILOCZYNÓW", MEDIAN: "MEDIANA", SIGN: "ZNAK.LICZBY", PI: "PI", EXP: "EXP", LN: "LN",
    LOG: "LOG", LOG10: "LOG10", CEILING: "ZAOKR.W.GÓRĘ", FLOOR: "ZAOKR.W.DÓŁ", "CEILING.MATH": "ZAOKR.W.GÓRĘ.MATEMATYCZNE",
    "FLOOR.MATH": "ZAOKR.W.DÓŁ.MATEMATYCZNE", MROUND: "ZAOKR.DO.WIELOKR", EVEN: "ZAOKR.DO.PARZ", ODD: "ZAOKR.DO.NPARZ",
    RAND: "LOS", RANDBETWEEN: "LOS.ZAKR", LARGE: "MAX.K", SMALL: "MIN.K", RANK: "POZYCJA", "RANK.EQ": "POZYCJA.NAJW",
    STDEV: "ODCH.STANDARDOWE", "STDEV.S": "ODCH.STANDARD.PRÓBKI", "STDEV.P": "ODCH.STANDARD.POPUL", VAR: "WARIANCJA",
    "VAR.S": "WARIANCJA.PRÓBKI", "VAR.P": "WARIANCJA.POPUL", SUBTOTAL: "SUMY.CZĘŚCIOWE", AVERAGEA: "ŚREDNIA.A",
    VLOOKUP: "WYSZUKAJ.PIONOWO", HLOOKUP: "WYSZUKAJ.POZIOMO", XLOOKUP: "X.WYSZUKAJ", LOOKUP: "WYSZUKAJ",
    INDEX: "INDEKS", MATCH: "PODAJ.POZYCJĘ", XMATCH: "X.DOPASUJ", ROW: "WIERSZ", COLUMN: "NR.KOLUMNY",
    ROWS: "ILE.WIERSZY", COLUMNS: "LICZBA.KOLUMN",
    TODAY: "DZIŚ", NOW: "TERAZ", DATE: "DATA", YEAR: "ROK", MONTH: "MIESIĄC", DAY: "DZIEŃ", WEEKDAY: "DZIEŃ.TYG",
    EDATE: "NR.SER.DATY", EOMONTH: "NR.SER.OST.DN.MIES", DATEDIF: "DATA.RÓŻNICA", DAYS: "DNI",
    NETWORKDAYS: "DNI.ROBOCZE", WORKDAY: "DZIEŃ.ROBOCZY", HOUR: "GODZINA", MINUTE: "MINUTA", SECOND: "SEKUNDA",
    TIME: "CZAS", WEEKNUM: "NUM.TYG", ISOWEEKNUM: "ISO.NUM.TYG", DATEVALUE: "DATA.WARTOŚĆ",
    LEN: "DŁ", LEFT: "LEWY", RIGHT: "PRAWY", MID: "FRAGMENT.TEKSTU", UPPER: "LITERY.WIELKIE", LOWER: "LITERY.MAŁE",
    PROPER: "Z.WIELKIEJ.LITERY", TRIM: "USUŃ.ZBĘDNE.ODSTĘPY", CLEAN: "OCZYŚĆ", CONCATENATE: "ZŁĄCZ.TEKSTY",
    CONCAT: "ZŁĄCZ.TEKST", TEXTJOIN: "POŁĄCZ.TEKSTY", TEXT: "TEKST", VALUE: "WARTOŚĆ", FIND: "ZNAJDŹ",
    SEARCH: "SZUKAJ.TEKST", SUBSTITUTE: "PODSTAW", REPLACE: "ZASTĄP", REPT: "POWT", EXACT: "PORÓWNAJ",
    CHAR: "ZNAK", CODE: "KOD", T: "T", N: "N",
    ISBLANK: "CZY.PUSTA", ISNUMBER: "CZY.LICZBA", ISTEXT: "CZY.TEKST", ISNONTEXT: "CZY.NIE.TEKST",
    ISLOGICAL: "CZY.LOGICZNA", ISERROR: "CZY.BŁĄD", ISERR: "CZY.BŁ", ISNA: "CZY.BRAK", ISEVEN: "CZY.PARZYSTE",
    ISODD: "CZY.NIEPARZYSTE", NA: "BRAK",
  };
  const EN_FROM_PL = {};
  Object.entries(PL_NAMES).forEach(([en, pl]) => { EN_FROM_PL[pl] = en; });
  // aliasy bez polskich znaków (klawiatura tabletu) i stare nazwy
  Object.entries(PL_NAMES).forEach(([en, pl]) => {
    const ascii = pl.normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/Ł/g, "L");
    if (!EN_FROM_PL[ascii]) EN_FROM_PL[ascii] = en;
  });
  // Funkcje dodane po Excelu 2007 — w pliku zapisuje się je z prefiksem _xlfn.
  const XLFN = new Set(["IFS", "IFNA", "XOR", "SWITCH", "XLOOKUP", "XMATCH", "CONCAT", "TEXTJOIN", "MAXIFS", "MINIFS",
    "DAYS", "CEILING.MATH", "FLOOR.MATH", "STDEV.S", "STDEV.P", "VAR.S", "VAR.P", "RANK.EQ", "ISOWEEKNUM"]);

  // ── tokenizer ──────────────────────────────────────────────────────────────
  const LETTER = "A-Za-z_\\u00C0-\\u024F";
  const RE_SHEET_QUOTED = /^'((?:[^']|'')+)'!/;
  const RE_SHEET_PLAIN = new RegExp(`^([${LETTER}][${LETTER}0-9.]*)!`);
  const RE_CELL = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?![A-Za-z0-9_(.])/;
  const RE_CELL_RANGE = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7}):(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?![A-Za-z0-9_(.])/;
  const RE_COL_RANGE = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![A-Za-z0-9_(.])/;
  const RE_ROW_RANGE = /^(\$?)(\d{1,7}):(\$?)(\d{1,7})(?![0-9.,])/;
  const RE_IDENT = new RegExp(`^[${LETTER}\\\\][${LETTER}0-9.\\\\]*`);
  const RE_ERROR = /^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|SPILL!|CALC!|ZERO!|DZIEL\/0!|ARG!|ADR!|NAZWA\?|LICZBA!|N\/D|ROZLANIE!|OBL!)/i;

  function colToIndex(letters) {
    let c = 0;
    for (const ch of letters.toUpperCase()) c = c * 26 + (ch.charCodeAt(0) - 64);
    return c - 1;
  }
  function indexToCol(i) {
    let s = "";
    let n = i + 1;
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }

  function syntaxError(msg, pos) {
    const e = new Error(msg);
    e.formulaSyntax = true;
    e.pos = pos;
    return e;
  }

  // Treść w nawiasach tabeli: Tabela[[#This Row],[od]] — z zagnieżdżonymi [] i ucieczką '.
  function readBracket(src, i) {
    let depth = 0;
    let j = i;
    for (; j < src.length; j++) {
      const ch = src[j];
      if (ch === "'" && j + 1 < src.length) { j++; continue; }
      if (ch === "[") depth++;
      else if (ch === "]") { depth--; if (depth === 0) return j + 1; }
    }
    throw syntaxError("Niedomknięty nawias [", i);
  }

  function tokenize(src, opts = {}) {
    const pl = opts.lang === "pl";
    const toks = [];
    let i = 0;
    let braces = 0; // wewnątrz stałej tablicowej „\” oddziela wiersze (PL), nie zaczyna nazwy
    const s = String(src);
    while (i < s.length) {
      const ch = s[i];
      if (ch === "{") braces++;
      else if (ch === "}") braces = Math.max(0, braces - 1);
      else if (ch === "\\" && braces) { toks.push({ t: "op", v: "\\", pos: i }); i++; continue; }
      if (ch === " " || ch === "\n" || ch === "\r" || ch === "\t" || ch === " ") { i++; continue; }
      const rest = s.slice(i);
      if (ch === '"') {
        let j = i + 1;
        let out = "";
        for (;;) {
          if (j >= s.length) throw syntaxError("Niedomknięty cudzysłów", i);
          if (s[j] === '"') {
            if (s[j + 1] === '"') { out += '"'; j += 2; continue; }
            break;
          }
          out += s[j++];
        }
        toks.push({ t: "str", v: out, pos: i });
        i = j + 1;
        continue;
      }
      const em = rest.match(RE_ERROR);
      if (em) {
        const up = em[0].toUpperCase();
        toks.push({ t: "err", v: ERR_FROM_PL[up] || up, pos: i });
        i += em[0].length;
        continue;
      }
      // odwołanie z nazwą arkusza
      let sheet = null;
      let sheetLen = 0;
      const sq = rest.match(RE_SHEET_QUOTED);
      const sp = !sq && rest.match(RE_SHEET_PLAIN);
      if (sq) { sheet = sq[1].replace(/''/g, "'"); sheetLen = sq[0].length; }
      else if (sp) { sheet = sp[1]; sheetLen = sp[0].length; }
      const body = sheetLen ? rest.slice(sheetLen) : rest;
      let m;
      if ((m = body.match(RE_CELL_RANGE))) {
        toks.push({ t: "range", pos: i, v: { sheet, r1: +m[4] - 1, c1: colToIndex(m[2]), r2: +m[8] - 1, c2: colToIndex(m[6]), ar1: !!m[3], ac1: !!m[1], ar2: !!m[7], ac2: !!m[5] } });
        i += sheetLen + m[0].length;
        continue;
      }
      if ((m = body.match(RE_CELL)) && !(body[m[0].length] === "(")) {
        toks.push({ t: "ref", pos: i, v: { sheet, r: +m[4] - 1, c: colToIndex(m[2]), ar: !!m[3], ac: !!m[1] } });
        i += sheetLen + m[0].length;
        continue;
      }
      if ((m = body.match(RE_COL_RANGE))) {
        toks.push({ t: "range", pos: i, v: { sheet, r1: 0, c1: colToIndex(m[2]), r2: Infinity, c2: colToIndex(m[4]), ac1: !!m[1], ac2: !!m[3], cols: true } });
        i += sheetLen + m[0].length;
        continue;
      }
      if (sheetLen && (m = body.match(RE_ROW_RANGE))) {
        toks.push({ t: "range", pos: i, v: { sheet, r1: +m[2] - 1, c1: 0, r2: +m[4] - 1, c2: Infinity, ar1: !!m[1], ar2: !!m[3], rows: true } });
        i += sheetLen + m[0].length;
        continue;
      }
      if (sheetLen) {
        // Arkusz!nazwa — nazwa lokalna arkusza
        const id = body.match(RE_IDENT);
        if (id) { toks.push({ t: "name", v: id[0], sheet, pos: i }); i += sheetLen + id[0].length; continue; }
        throw syntaxError("Po „!” oczekiwano adresu komórki", i);
      }
      // liczby (PL: przecinek dziesiętny, gdy po nim cyfry)
      const numRe = pl ? /^(\d+(?:[.,]\d+)?|[.,]\d+)(?:[eE][+-]?\d+)?/ : /^(\d+(?:\.\d+)?|\.\d+)(?:[eE][+-]?\d+)?/;
      if ((m = rest.match(numRe)) && !(/^\d+:/.test(rest) && rest.match(RE_ROW_RANGE))) {
        toks.push({ t: "num", v: Number(m[0].replace(",", ".")), raw: m[0], pos: i });
        i += m[0].length;
        continue;
      }
      if ((m = rest.match(RE_ROW_RANGE))) {
        toks.push({ t: "range", pos: i, v: { sheet: null, r1: +m[2] - 1, c1: 0, r2: +m[4] - 1, c2: Infinity, ar1: !!m[1], ar2: !!m[3], rows: true } });
        i += m[0].length;
        continue;
      }
      const id = rest.match(RE_IDENT);
      if (id) {
        const word = id[0];
        let j = i + word.length;
        if (s[j] === "[") {
          const end = readBracket(s, j);
          toks.push({ t: "sref", v: { table: word, spec: s.slice(j + 1, end - 1) }, pos: i });
          i = end;
          continue;
        }
        let k = j;
        while (s[k] === " ") k++;
        if (s[k] === "(") {
          toks.push({ t: "fn", v: word, pos: i });
          i = k + 1;
          continue;
        }
        const up = word.toUpperCase();
        if (up === "TRUE" || up === "PRAWDA") toks.push({ t: "bool", v: true, pos: i });
        else if (up === "FALSE" || up === "FAŁSZ" || up === "FALSZ") toks.push({ t: "bool", v: false, pos: i });
        else toks.push({ t: "name", v: word, sheet: null, pos: i });
        i = j;
        continue;
      }
      const two = s.slice(i, i + 2);
      if (two === "<>" || two === "<=" || two === ">=") { toks.push({ t: "op", v: two, pos: i }); i += 2; continue; }
      if ("+-*/^&=<>%(),;{}".includes(ch)) {
        // „;” w zapisie angielskim/pliku = separator WIERSZY stałej tablicowej (semi); w PL to
        // separator argumentów i kolumn, a wiersze oddziela „\\” (tak pisze serialize)
        if (ch === ";") toks.push({ t: "op", v: ",", pos: i, semi: !pl });
        else toks.push({ t: "op", v: ch, pos: i });
        i++;
        continue;
      }
      throw syntaxError(`Nieznany znak „${ch}”`, i);
    }
    return toks;
  }

  // ── parser ─────────────────────────────────────────────────────────────────
  const CMP = new Set(["=", "<>", "<", ">", "<=", ">="]);

  function canonicalFn(name) {
    let up = String(name).toUpperCase();
    up = up.replace(/^_XLFN\./, "").replace(/^_XLWS\./, "").replace(/^_XLFN\._XLWS\./, "");
    if (EN_FROM_PL[up]) return EN_FROM_PL[up];
    return up;
  }

  // Polski zapis (przecinek dziesiętny, „;”) czy angielski (kropka, „,”)? W języku PL:
  // średnik albo polska nazwa funkcji → polski; same angielskie nazwy bez średnika
  // (np. wklejone „IF(A1>0,1,0)”) → angielski; bez funkcji (=A1*1,5) → polski.
  function syntaxLang(text, lang) {
    if (lang !== "pl") return "en";
    const bare = String(text).replace(/"(?:[^"]|"")*"/g, '""').replace(/'(?:[^']|'')*'!/g, "S!");
    if (bare.includes(";")) return "pl";
    const names = bare.match(/[A-Za-z_\u00C0-\u024F][A-Za-z0-9_.\u00C0-\u024F]*(?=\s*\()/g) || [];
    let en = false;
    for (const raw of names) {
      const up = raw.toUpperCase().replace(/^_XLFN\./, "");
      if (EN_FROM_PL[up] && PL_NAMES[EN_FROM_PL[up]] && EN_FROM_PL[up] !== up) return "pl";
      if (PL_NAMES[up] && PL_NAMES[up] !== up) en = true;
    }
    return en ? "en" : "pl";
  }

  function parse(src, opts = {}) {
    const text = String(src || "").replace(/^\s*=/, "");
    const toks = tokenize(text, { ...opts, lang: syntaxLang(text, opts.lang) });
    let p = 0;
    const peek = () => toks[p];
    const isOp = (v) => toks[p] && toks[p].t === "op" && toks[p].v === v;
    const expectOp = (v) => {
      if (!isOp(v)) throw syntaxError(v === ")" ? "Brakuje nawiasu zamykającego „)”" : `Oczekiwano „${v}”`, toks[p] ? toks[p].pos : text.length);
      p++;
    };

    function comparison() {
      let left = concat();
      while (toks[p] && toks[p].t === "op" && CMP.has(toks[p].v)) {
        const op = toks[p++].v;
        left = { t: "bin", op, a: left, b: concat() };
      }
      return left;
    }
    function concat() {
      let left = additive();
      while (isOp("&")) { p++; left = { t: "bin", op: "&", a: left, b: additive() }; }
      return left;
    }
    function additive() {
      let left = multiplicative();
      while (isOp("+") || isOp("-")) { const op = toks[p++].v; left = { t: "bin", op, a: left, b: multiplicative() }; }
      return left;
    }
    function multiplicative() {
      let left = power();
      while (isOp("*") || isOp("/")) { const op = toks[p++].v; left = { t: "bin", op, a: left, b: power() }; }
      return left;
    }
    function power() {
      let left = unary();
      while (isOp("^")) { p++; left = { t: "bin", op: "^", a: left, b: unary() }; }
      return left;
    }
    function unary() {
      if (isOp("-")) { p++; return { t: "un", op: "-", a: unary() }; }
      if (isOp("+")) { p++; return { t: "un", op: "+", a: unary() }; }
      return postfix();
    }
    function postfix() {
      let node = primary();
      while (isOp("%")) { p++; node = { t: "pct", a: node }; }
      return node;
    }
    function primary() {
      const tok = peek();
      if (!tok) throw syntaxError("Formuła urywa się — czegoś brakuje na końcu", text.length);
      if (tok.t === "num") { p++; return { t: "num", v: tok.v }; }
      if (tok.t === "str") { p++; return { t: "str", v: tok.v }; }
      if (tok.t === "bool") { p++; return { t: "bool", v: tok.v }; }
      if (tok.t === "err") { p++; return { t: "err", v: tok.v }; }
      if (tok.t === "ref") { p++; return { t: "ref", ...tok.v }; }
      if (tok.t === "range") { p++; return { t: "range", ...tok.v }; }
      if (tok.t === "sref") { p++; return { t: "sref", ...tok.v }; }
      if (tok.t === "name") { p++; return { t: "name", v: tok.v, sheet: tok.sheet }; }
      if (tok.t === "fn") {
        p++;
        const name = canonicalFn(tok.v);
        const args = [];
        if (isOp(")")) { p++; return { t: "call", name, raw: tok.v, args }; }
        for (;;) {
          if (isOp(",")) { args.push({ t: "missing" }); p++; continue; }
          if (isOp(")")) { args.push({ t: "missing" }); p++; break; }
          args.push(comparison());
          if (isOp(",")) { p++; if (isOp(")")) { args.push({ t: "missing" }); p++; break; } continue; }
          expectOp(")");
          break;
        }
        return { t: "call", name, raw: tok.v, args };
      }
      if (tok.t === "op" && tok.v === "(") {
        p++;
        const inner = comparison();
        expectOp(")");
        return { t: "paren", a: inner };
      }
      if (tok.t === "op" && tok.v === "{") {
        // stała tablicowa {1,2;3,4} (EN/plik) albo {1;2\3;4} (PL): kolumny , / ;  wiersze ; / \
        // Dawniej „;” w EN brane było za kolumnę — {1,"a";2,"b"} spłaszczało się do jednego
        // wiersza (WYSZUKAJ.PIONOWO → #N/D!) i tak zostawało w pliku po zapisie.
        p++;
        const rows = [[]];
        for (;;) {
          const v = unary();
          rows[rows.length - 1].push(v);
          if ((isOp(",") && toks[p].semi) || isOp("\\")) { p++; rows.push([]); continue; }
          if (isOp(",")) { p++; continue; }
          if (isOp("}")) { p++; break; }
          throw syntaxError("Niepoprawna stała tablicowa", tok.pos);
        }
        return { t: "array", rows };
      }
      throw syntaxError(`Nieoczekiwane „${tok.v ?? tok.t}”`, tok.pos);
    }

    if (!toks.length) throw syntaxError("Pusta formuła", 0);
    const ast = comparison();
    if (p < toks.length) {
      const tok = toks[p];
      if (tok.t === "op" && tok.v === ")") throw syntaxError("Za dużo nawiasów zamykających „)”", tok.pos);
      throw syntaxError(`Nieoczekiwane „${tok.v ?? tok.t}”`, tok.pos);
    }
    return ast;
  }

  // ── zapis z powrotem do tekstu ─────────────────────────────────────────────
  function quoteSheet(name) {
    if (/^[A-Za-z_À-ɏ][A-Za-z0-9_.À-ɏ]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) && !/^(TRUE|FALSE)$/i.test(name)) return name;
    return `'${String(name).replace(/'/g, "''")}'`;
  }
  function refText(r, c, ar, ac) {
    return `${ac ? "$" : ""}${indexToCol(c)}${ar ? "$" : ""}${r + 1}`;
  }
  function numText(n, pl) {
    if (!Number.isFinite(n)) return "0";
    let s = String(n);
    if (/e/i.test(s)) s = s.toUpperCase().replace(/E\+?(-?)(\d)$/, "E$10$2").replace("E+", "E+");
    return pl ? s.replace(".", ",") : s;
  }

  // opts.lang: "pl" (nazwy PL, „;”, przecinek) albo "en"; opts.file: zapis do pliku (+ _xlfn.)
  function serialize(ast, opts = {}) {
    const pl = opts.lang === "pl" && !opts.file;
    const sep = pl ? ";" : ",";
    const out = (n) => {
      switch (n.t) {
        case "num": return numText(n.v, pl);
        case "str": return `"${String(n.v).replace(/"/g, '""')}"`;
        case "bool": return pl ? (n.v ? "PRAWDA" : "FAŁSZ") : (n.v ? "TRUE" : "FALSE");
        case "err": return pl ? (ERR_PL[n.v] || n.v) : n.v;
        case "missing": return "";
        case "ref": return (n.sheet != null ? quoteSheet(n.sheet) + "!" : "") + refText(n.r, n.c, n.ar, n.ac);
        case "range": {
          const pre = n.sheet != null ? quoteSheet(n.sheet) + "!" : "";
          if (n.cols) return `${pre}${n.ac1 ? "$" : ""}${indexToCol(n.c1)}:${n.ac2 ? "$" : ""}${indexToCol(n.c2)}`;
          if (n.rows) return `${pre}${n.ar1 ? "$" : ""}${n.r1 + 1}:${n.ar2 ? "$" : ""}${n.r2 + 1}`;
          return `${pre}${refText(n.r1, n.c1, n.ar1, n.ac1)}:${refText(n.r2, n.c2, n.ar2, n.ac2)}`;
        }
        case "sref": return `${n.table}[${n.spec}]`;
        case "name": return (n.sheet != null ? quoteSheet(n.sheet) + "!" : "") + n.v;
        case "paren": return `(${out(n.a)})`;
        case "un": return n.op + out(n.a);
        case "pct": return out(n.a) + "%";
        case "bin": return out(n.a) + n.op + out(n.b);
        case "array": return "{" + n.rows.map((r) => r.map(out).join(sep)).join(pl ? "\\" : ";") + "}";
        case "call": {
          let name = n.name;
          if (pl) name = PL_NAMES[n.name] || n.name;
          else if (opts.file && XLFN.has(n.name)) name = "_xlfn." + n.name;
          return `${name}(${n.args.map(out).join(sep)})`;
        }
        default: return "";
      }
    };
    return out(ast);
  }

  // Tekst formuły z pliku → do pokazania w języku aplikacji (bez „=”).
  function localize(fileFormula, lang) {
    try { return serialize(parse(fileFormula, { lang: "en" }), { lang }); } catch { return String(fileFormula || "").replace(/_xlfn\./gi, ""); }
  }

  // Przesunięcie odwołań względnych (wypełnianie w dół / w prawo, jak w Excelu).
  // Odwołanie, które wypadłoby poza arkusz, daje #ADR! (jak Excel).
  function shift(ast, dr, dc) {
    const walk = (n) => {
      if (!n || typeof n !== "object") return n;
      if (n.t === "ref") {
        const r = n.ar ? n.r : n.r + dr;
        const c = n.ac ? n.c : n.c + dc;
        if (r < 0 || c < 0) return { t: "err", v: "#REF!" };
        return { ...n, r, c };
      }
      if (n.t === "range") {
        const m = { ...n };
        if (!n.cols) { if (!n.ar1) m.r1 = n.r1 + dr; if (!n.ar2) m.r2 = n.r2 + dr; }
        if (!n.rows) { if (!n.ac1) m.c1 = n.c1 + dc; if (!n.ac2) m.c2 = n.c2 + dc; }
        if (m.r1 < 0 || m.c1 < 0 || m.r2 < 0 || m.c2 < 0) return { t: "err", v: "#REF!" };
        return m;
      }
      if (n.t === "call") return { ...n, args: n.args.map(walk) };
      if (n.t === "bin") return { ...n, a: walk(n.a), b: walk(n.b) };
      if (n.t === "un" || n.t === "pct" || n.t === "paren") return { ...n, a: walk(n.a) };
      if (n.t === "array") return { ...n, rows: n.rows.map((r) => r.map(walk)) };
      return n;
    };
    return walk(ast);
  }

  // ── wartości ───────────────────────────────────────────────────────────────
  // Zakres jako obiekt (nie tablica) — funkcje typu LICZ.JEŻELI/INDEKS/WIERSZ potrzebują
  // wiedzieć, SKĄD pochodzą komórki; do tablicy materializujemy dopiero przy działaniu.
  class Range {
    constructor(ctx, sheet, r1, c1, r2, c2) {
      this.ctx = ctx; this.sheet = sheet;
      this.r1 = Math.min(r1, r2); this.r2 = Math.max(r1, r2);
      this.c1 = Math.min(c1, c2); this.c2 = Math.max(c1, c2);
      if (!Number.isFinite(this.r2) || !Number.isFinite(this.c2)) {
        const b = ctx.bounds(sheet);
        if (!Number.isFinite(this.r2)) this.r2 = Math.max(this.r1, b.r2);
        if (!Number.isFinite(this.c2)) this.c2 = Math.max(this.c1, b.c2);
      }
    }
    get rows() { return this.r2 - this.r1 + 1; }
    get cols() { return this.c2 - this.c1 + 1; }
    at(i, j) { return this.ctx.get(this.sheet, this.r1 + i, this.c1 + j); }
    // ctx.matrixCache (makeMatrixCache) — przeliczenie wielu formuł naraz: ten sam zakres
    // (np. cała kolumna pod LICZ.JEŻELI w 300 formułach) budujemy RAZ, nie 300 razy.
    // Wynik jest współdzielony, więc nikt nie może go modyfikować (funkcje tylko czytają).
    toMatrix() {
      const cache = this.ctx.matrixCache;
      const key = cache ? `${this.sheet}\u0000${this.r1}\u0000${this.c1}\u0000${this.r2}\u0000${this.c2}` : null;
      if (cache) { const hit = cache.get(key); if (hit) return hit; }
      const out = [];
      for (let i = 0; i < this.rows; i++) {
        const row = [];
        for (let j = 0; j < this.cols; j++) row.push(this.at(i, j));
        out.push(row);
      }
      if (cache) cache.set(key, out, this);
      return out;
    }
    *values() {
      if (this.ctx.matrixCache) { for (const row of this.toMatrix()) yield* row; return; }
      for (let i = 0; i < this.rows; i++) for (let j = 0; j < this.cols; j++) yield this.at(i, j);
    }
  }
  // Pamięć zbudowanych zakresów na czas jednego przeliczenia. invalidate(arkusz, r, c) —
  // komórka się zmieniła (np. wynik formuły w tym przebiegu) → zakresy z nią wypadają.
  function makeMatrixCache() {
    const map = new Map();
    return {
      get(key) { const e = map.get(key); return e ? e.m : null; },
      set(key, m, rg) { map.set(key, { m, sheet: rg.sheet, r1: rg.r1, c1: rg.c1, r2: rg.r2, c2: rg.c2 }); },
      invalidate(sheet, r, c) {
        for (const [key, e] of map) if (e.sheet === sheet && r >= e.r1 && r <= e.r2 && c >= e.c1 && c <= e.c2) map.delete(key);
      },
      get size() { return map.size; },
    };
  }
  const isRange = (v) => v instanceof Range;
  const isMatrix = (v) => Array.isArray(v) && Array.isArray(v[0]);
  const toMatrix = (v) => (isRange(v) ? v.toMatrix() : isMatrix(v) ? v : [[v]]);
  function* flat(v) {
    if (isRange(v)) yield* v.values();
    else if (isMatrix(v)) for (const r of v) yield* r;
    else yield v;
  }
  // Wartość skalarna z zakresu/tablicy w miejscu, gdzie oczekiwana jest jedna wartość
  // (niejawne przecięcie Excela: kolumna → wiersz formuły, wiersz → kolumna formuły).
  function scalar(v, ctx) {
    if (isRange(v)) {
      if (v.rows === 1 && v.cols === 1) return v.at(0, 0);
      if (ctx && v.cols === 1 && ctx.row >= v.r1 && ctx.row <= v.r2 && ctx.sheetName === v.sheet) return v.at(ctx.row - v.r1, 0);
      if (ctx && v.rows === 1 && ctx.col >= v.c1 && ctx.col <= v.c2 && ctx.sheetName === v.sheet) return v.at(0, ctx.col - v.c1);
      return v.at(0, 0);
    }
    if (isMatrix(v)) return v[0][0];
    return v;
  }

  // Tekst → liczba w działaniach, jak Excel w danych ustawieniach regionalnych.
  // en: "1,000.5", "$1,000", "(5)" = -5, "50%", "1e3", "12:00", "2:15 PM";
  // pl: "1 234,5", "12,5 zł", "50%". Daty ISO i dd.mm.rrrr w obu. Reszta → null (#ARG!).
  let numberLocale = "en";
  function setLocale(lang) { numberLocale = lang === "pl" ? "pl" : "en"; }
  function parseTimeText(t) {
    const m = t.match(/^(\d{1,2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?\s*(AM|PM|am|pm)?$/);
    if (!m) return null;
    let h = +m[1];
    const mi = +m[2], se = m[3] ? +m[3] : 0;
    if (m[4]) { if (h > 12) return null; h = (h % 12) + (/p/i.test(m[4]) ? 12 : 0); }
    if (mi > 59 || se >= 60) return null;
    return (h * 3600 + mi * 60 + se) / 86400;
  }
  function parseNumberText(s) {
    let t = String(s).replace(/\u00A0/g, " ").trim();
    if (!t) return null;
    let m = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (m) return dateSerial(+m[1], +m[2], +m[3]);
    m = t.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
    if (m) return dateSerial(+m[3], +m[2], +m[1]);
    const time = parseTimeText(t);
    if (time !== null) return time;
    let neg = false;
    if (/^\(.*\)$/.test(t)) { neg = true; t = t.slice(1, -1).trim(); }
    let pct = 0;
    while (t.endsWith("%")) { pct++; t = t.slice(0, -1).trim(); }
    if (/^[+-]/.test(t)) { if (t[0] === "-") neg = !neg; t = t.slice(1).trim(); }
    if (numberLocale === "pl") {
      t = t.replace(/\s*zł$/i, "").replace(/^zł\s*/i, "");
      if (!/^(?:\d{1,3}(?: \d{3})+|\d+)?(?:[,.]\d+)?(?:[eE][+-]?\d+)?$/.test(t) || !/\d/.test(t)) return null;
      t = t.replace(/ /g, "").replace(",", ".");
    } else {
      t = t.replace(/^\$\s*/, "");
      if (!/^[\d,]*(?:\.\d*)?(?:[eE][+-]?\d+)?$/.test(t) || !/\d/.test(t) || /^,/.test(t)) return null;
      t = t.replace(/,/g, "");
    }
    let n = Number(t);
    if (!Number.isFinite(n)) return null;
    for (let i = 0; i < pct; i++) n /= 100;
    return neg ? -n : n;
  }

  // ── arytmetyka jak w Excelu ────────────────────────────────────────────────
  // Wykładnik binarny liczby (jak frexp − 1): 2^e ≤ |x| < 2^(e+1).
  const _dv = new DataView(new ArrayBuffer(8));
  function binExp(x) {
    _dv.setFloat64(0, x);
    const e = (_dv.getUint32(0) >>> 20) & 0x7ff;
    return e === 0 ? -1074 : e - 1023;
  }
  // Excel zeruje wynik dodawania/odejmowania bliski zera: |wynik| < 2^(e−49), e = wykładnik
  // LEWEGO składnika — tylko w korzeniu formuły i przy ostatnim dodawaniu SUMY.
  function snap(left, result) {
    if (typeof left !== "number" || typeof result !== "number" || result === 0 || left === 0) return result;
    return Math.abs(result) < Math.pow(2, binExp(left) - 49) ? 0 : result;
  }
  // Liczba jako 15 cyfr znaczących (tak Excel ją „widzi” w ZAOKR): { neg, d: "15 cyfr", x: wykładnik }.
  // Remis na 16. cyfrze idzie w stronę zera (ZAOKR(1234567890123,375;2) = …,37).
  function decimal15(v) {
    const a = Math.abs(v);
    const exact = a.toExponential(99);
    const m = exact.match(/^(\d)\.(\d+)e([+-]\d+)$/);
    const digits = m[1] + m[2];
    let x = Number(m[3]);
    const tail = digits.slice(15);
    let d = digits.slice(0, 15);
    if (tail[0] > "5" || (tail[0] === "5" && /[1-9]/.test(tail.slice(1)))) {
      const up = (BigInt(d) + 1n).toString();
      if (up.length > 15) { d = up.slice(0, 15); x += 1; } else d = up;
    }
    return { neg: v < 0, d, x };
  }
  function roundExcel(v, digits, mode) {
    if (v === 0 || !Number.isFinite(v)) return v;
    let k = Number(Number(digits).toPrecision(15));
    k = k < 0 ? Math.ceil(k) : Math.floor(k);
    if (k > 300) return v;
    const { neg, d, x } = decimal15(v);
    const keep = x + k + 1; // ile cyfr z 15 zostaje
    const sign = neg ? -1 : 1;
    if (keep >= 15) return sign * Number(`${d[0]}.${d.slice(1)}e${x}`) || 0;
    let int;
    if (keep <= 0) {
      const rest = keep === 0 ? d : "";
      const up = mode === "half" ? keep === 0 && d[0] >= "5" : mode === "up";
      if (!up) return 0;
      int = "1";
      if (keep < 0 && mode !== "up") return 0;
      return (sign * Number(`1e${x - keep + 1}`)) || 0;
    }
    int = d.slice(0, keep);
    const rest = d.slice(keep);
    const up = mode === "half" ? rest[0] >= "5" : mode === "up" ? /[1-9]/.test(rest) : false;
    if (up) int = (BigInt(int) + 1n).toString();
    return (sign * Number(`${int}e${x - keep + 1}`)) || 0;
  }
  // Potęga jak w Excelu: całkowity wykładnik = kolejne podnoszenie do kwadratu (bit w bit),
  // x^−n = 1/x^n; ułamkowy: exp(b·ln x), pierwiastek przy 0,5; ujemna podstawa tylko
  // przy 1/b nieparzystym całkowitym. 0^0 = #LICZBA!, 0^−1 = #DZIEL/0!.
  function excelPow(x, y) {
    if (x === 0) {
      if (y === 0) return ERR["#NUM!"];
      if (y < 0) return ERR["#DIV/0!"];
      return 0;
    }
    let r;
    if (Number.isInteger(y)) {
      let n = Math.abs(y), b = x;
      r = 1;
      while (n > 0) {
        if (n % 2 === 1) r *= b;
        n = Math.floor(n / 2);
        if (n > 0) b *= b;
      }
      if (y < 0) r = 1 / r;
    } else if (x > 0) {
      if (y === 0.5) r = Math.sqrt(x);
      else if (y < 0) r = 1 / Math.exp(-y * Math.log(x));
      else r = Math.exp(y * Math.log(x));
    } else {
      const inv = 1 / y;
      if (!Number.isInteger(inv) || Math.abs(inv) % 2 !== 1) return ERR["#NUM!"];
      const m = Math.abs(x);
      r = y < 0 ? -1 / Math.exp(-y * Math.log(m)) : -Math.exp(y * Math.log(m));
    }
    return Number.isFinite(r) ? r : ERR["#NUM!"];
  }

  function num(v) {
    if (typeof v === "number") return Math.abs(v) < 2.2250738585072014e-308 ? 0 : v; // podnormalna = 0
    if (v === null || v === undefined || v === "") return v === "" ? ERR["#VALUE!"] : 0;
    if (typeof v === "boolean") return v ? 1 : 0;
    if (isErr(v)) return v;
    if (typeof v === "string") { const n = parseNumberText(v); return n === null ? ERR["#VALUE!"] : n; }
    return ERR["#VALUE!"];
  }
  function generalText(n) {
    if (!Number.isFinite(n)) return "#NUM!";
    if (Number.isInteger(n) && Math.abs(n) < 1e15) return String(n);
    const a = Math.abs(n);
    if (a !== 0 && (a < 1e-9 || a >= 1e15)) return n.toExponential(5).replace(/\.?0+e/, "E").replace("e", "E").replace(/E\+?(-?)(\d)$/, "E+$10$2").replace("E+-", "E-");
    return String(Number(n.toPrecision(15)));
  }
  function text(v, ctx) {
    if (typeof v === "string") return v;
    if (v === null || v === undefined) return "";
    if (typeof v === "boolean") return ctx && ctx.lang === "pl" ? (v ? "PRAWDA" : "FAŁSZ") : (v ? "TRUE" : "FALSE");
    if (typeof v === "number") { const s = generalText(v); return ctx && ctx.lang === "pl" ? s.replace(".", ",") : s; }
    if (isErr(v)) return v;
    return String(v);
  }
  function bool(v) {
    if (typeof v === "boolean") return v;
    if (typeof v === "number") return v !== 0;
    if (v === null || v === undefined) return false;
    if (isErr(v)) return v;
    if (typeof v === "string") {
      const u = v.toUpperCase();
      if (u === "TRUE" || u === "PRAWDA") return true;
      if (u === "FALSE" || u === "FAŁSZ") return false;
      return ERR["#VALUE!"];
    }
    return ERR["#VALUE!"];
  }
  const typeRank = (v) => (typeof v === "number" ? 1 : typeof v === "string" ? 2 : typeof v === "boolean" ? 3 : 0);
  const COLLATOR = typeof Intl !== "undefined" ? new Intl.Collator("pl", { sensitivity: "accent" }) : null;
  function textCompare(a, b) {
    const x = a.toLowerCase(), y = b.toLowerCase();
    if (x === y) return 0;
    return COLLATOR ? COLLATOR.compare(x, y) : (x < y ? -1 : 1);
  }
  // −1 / 0 / 1 jak w Excelu (puste = 0 albo "" zależnie od drugiej strony)
  function compare(a, b) {
    if (typeof a === "number" && Math.abs(a) < 2.2250738585072014e-308) a = 0; // podnormalna = 0
    if (typeof b === "number" && Math.abs(b) < 2.2250738585072014e-308) b = 0;
    if (a === null || a === undefined) a = typeof b === "string" ? "" : typeof b === "boolean" ? false : 0;
    if (b === null || b === undefined) b = typeof a === "string" ? "" : typeof a === "boolean" ? false : 0;
    const ta = typeRank(a), tb = typeRank(b);
    if (ta !== tb) return ta < tb ? -1 : 1;
    if (ta === 1) {
      if (a === b) return 0;
      // Excel porównuje liczby z dokładnością 15 cyfr znaczących (0,1+0,2 = 0,3 → PRAWDA)
      if (Number(a.toPrecision(15)) === Number(b.toPrecision(15))) return 0;
      return a < b ? -1 : 1;
    }
    if (ta === 2) return textCompare(a, b);
    if (ta === 3) return a === b ? 0 : (a ? 1 : -1);
    return 0;
  }

  // ── daty (system 1900, dni od 1899-12-30) ──────────────────────────────────
  const DAY_MS = 86400000;
  const EPOCH = Date.UTC(1899, 11, 30);
  function dateSerial(y, m, d) { return (Date.UTC(y, m - 1, d) - EPOCH) / DAY_MS; }
  function serialToParts(s) {
    const d = new Date(EPOCH + Math.floor(s) * DAY_MS);
    return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate(), wd: d.getUTCDay() };
  }
  function jsDateToSerial(d) {
    return (Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds(), d.getMilliseconds()) - EPOCH) / DAY_MS;
  }
  function daysInMonth(y, m) { return new Date(Date.UTC(y, m, 0)).getUTCDate(); }

  // ── kryteria (LICZ.JEŻELI, SUMA.WARUNKÓW…) ─────────────────────────────────
  function wildcardRe(pattern) {
    let re = "";
    for (let i = 0; i < pattern.length; i++) {
      const ch = pattern[i];
      if (ch === "~" && i + 1 < pattern.length) { re += pattern[++i].replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); continue; }
      if (ch === "*") re += "[\\s\\S]*";
      else if (ch === "?") re += "[\\s\\S]";
      else re += ch.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    }
    return new RegExp(`^${re}$`, "i");
  }
  function makeCriterion(crit) {
    if (isErr(crit)) return () => false;
    if (typeof crit === "number") return (v) => (typeof v === "number" ? v === crit : (typeof v === "string" && parseNumberText(v) === crit));
    if (typeof crit === "boolean") return (v) => v === crit;
    if (crit === null || crit === undefined) return (v) => v === null || v === "";
    const s = String(crit);
    const m = s.match(/^(<=|>=|<>|<|>|=)?([\s\S]*)$/);
    const op = m[1] || "=";
    const rhs = m[2];
    if (op === "=" && rhs === "" && m[1]) return (v) => v === null || v === "";
    if (op === "<>" && rhs === "") return (v) => !(v === null || v === "");
    if (!m[1] && rhs === "") return (v) => v === null || v === "";
    const n = parseNumberText(rhs);
    const b = /^(TRUE|PRAWDA)$/i.test(rhs) ? true : /^(FALSE|FAŁSZ)$/i.test(rhs) ? false : null;
    if (n !== null) {
      return (v) => {
        const x = typeof v === "number" ? v : (typeof v === "string" && (op === "=" || op === "<>") ? parseNumberText(v) : null);
        if (x === null || x === undefined) return op === "<>";
        switch (op) {
          case "=": return x === n; case "<>": return x !== n; case "<": return x < n;
          case ">": return x > n; case "<=": return x <= n; case ">=": return x >= n;
          default: return false;
        }
      };
    }
    if (b !== null) return (v) => (op === "<>" ? v !== b : op === "=" ? v === b : false);
    if (op === "=" || op === "<>") {
      const hasWild = /[*?~]/.test(rhs);
      const re = hasWild ? wildcardRe(rhs) : null;
      const low = rhs.toLowerCase();
      const test = (v) => {
        if (typeof v !== "string") return false;
        return re ? re.test(v) : v.toLowerCase() === low;
      };
      return op === "=" ? test : (v) => !test(v);
    }
    return (v) => {
      if (typeof v !== "string") return false;
      const c = textCompare(v, rhs);
      switch (op) { case "<": return c < 0; case ">": return c > 0; case "<=": return c <= 0; case ">=": return c >= 0; default: return false; }
    };
  }

  // ── funkcje ────────────────────────────────────────────────────────────────
  // Każda dostaje (args: wartości już policzone — zakresy jako Range, ctx).
  // Funkcje „leniwe” (JEŻELI, ORAZ…) dostają węzły AST i same je liczą.
  const FN = {};
  const LAZY = {};

  function numbersFrom(args, { countText = false } = {}) {
    const out = [];
    for (const a of args) {
      if (isRange(a) || isMatrix(a)) {
        for (const v of flat(a)) {
          if (isErr(v)) return v;
          if (typeof v === "number") out.push(v);
        }
      } else if (a !== undefined && !(a && a.t === "missing")) {
        if (a === null) { if (countText) out.push(0); continue; }
        const n = num(a);
        if (isErr(n)) return n;
        out.push(n);
      }
    }
    return out;
  }
  const agg = (fn) => (args) => { const ns = numbersFrom(args); return isErr(ns) ? ns : fn(ns); };
  FN.SUM = agg((ns) => {
    let total = 0, prev = 0;
    for (const n of ns) { prev = total; total += n; }
    if (!Number.isFinite(total)) return ERR["#NUM!"];
    return ns.length > 1 ? snap(prev, total) : total;
  });
  FN.PRODUCT = agg((ns) => (ns.length ? ns.reduce((a, b) => a * b, 1) : 0));
  FN.AVERAGE = agg((ns) => {
    if (!ns.length) return ERR["#DIV/0!"];
    const t = ns.reduce((a, b) => a + b, 0);
    return Number.isFinite(t) ? t / ns.length : ERR["#NUM!"];
  });
  // pętla zamiast Math.min(...ns): od ~125 tys. liczb (MAX(A:A) w długim arkuszu) Chromium rzuca RangeError
  const minOf = (ns) => { let m = Infinity; for (const x of ns) if (x < m) m = x; return m; };
  const maxOf = (ns) => { let m = -Infinity; for (const x of ns) if (x > m) m = x; return m; };
  FN.MIN = agg((ns) => (ns.length ? minOf(ns) : 0));
  FN.MAX = agg((ns) => (ns.length ? maxOf(ns) : 0));
  FN.MEDIAN = agg((ns) => {
    if (!ns.length) return ERR["#NUM!"];
    const s = ns.slice().sort((a, b) => a - b);
    const m = s.length >> 1;
    return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
  });
  const variance = (ns, pop) => {
    const n = ns.length;
    if (n < (pop ? 1 : 2)) return ERR["#DIV/0!"];
    const mean = ns.reduce((a, b) => a + b, 0) / n;
    return ns.reduce((a, b) => a + (b - mean) ** 2, 0) / (pop ? n : n - 1);
  };
  FN.VAR = FN["VAR.S"] = agg((ns) => variance(ns, false));
  FN["VAR.P"] = agg((ns) => variance(ns, true));
  FN.STDEV = FN["STDEV.S"] = agg((ns) => { const v = variance(ns, false); return isErr(v) ? v : Math.sqrt(v); });
  FN["STDEV.P"] = agg((ns) => { const v = variance(ns, true); return isErr(v) ? v : Math.sqrt(v); });
  FN.COUNT = (args) => {
    let n = 0;
    for (const a of args) {
      if (isRange(a) || isMatrix(a)) { for (const v of flat(a)) if (typeof v === "number") n++; }
      else if (typeof a === "number" || typeof a === "boolean" || (typeof a === "string" && parseNumberText(a) !== null)) n++;
    }
    return n;
  };
  FN.COUNTA = (args) => {
    let n = 0;
    for (const a of args) {
      if (isRange(a) || isMatrix(a)) { for (const v of flat(a)) if (v !== null && v !== undefined) n++; }
      else if (!(a && a.t === "missing")) n++;
    }
    return n;
  };
  FN.COUNTBLANK = (args) => {
    let n = 0;
    for (const v of flat(args[0])) if (v === null || v === "") n++;
    return n;
  };
  FN.AVERAGEA = (args) => {
    const ns = [];
    for (const a of args) for (const v of flat(a)) {
      if (isErr(v)) return v;
      if (typeof v === "number") ns.push(v);
      else if (typeof v === "boolean") ns.push(v ? 1 : 0);
      else if (typeof v === "string") ns.push(0);
    }
    return ns.length ? ns.reduce((a, b) => a + b, 0) / ns.length : ERR["#DIV/0!"];
  };

  function sameShape(a, b) {
    if (!isRange(a) || !isRange(b)) return false;
    return a.rows === b.rows && a.cols === b.cols;
  }
  // Pary (zakres, kryterium) → maska zgodnych pozycji.
  function criteriaMask(pairs) {
    const first = pairs[0][0];
    if (!isRange(first) && !isMatrix(first)) return ERR["#VALUE!"];
    const m0 = toMatrix(first);
    const rows = m0.length, cols = m0[0].length;
    const mask = new Uint8Array(rows * cols).fill(1);
    for (const [rng, crit] of pairs) {
      const mat = toMatrix(rng);
      if (mat.length !== rows || mat[0].length !== cols) return ERR["#VALUE!"];
      const test = makeCriterion(scalar(crit));
      for (let i = 0; i < rows; i++) for (let j = 0; j < cols; j++) {
        const k = i * cols + j;
        if (mask[k] && !test(mat[i][j])) mask[k] = 0;
      }
    }
    return { mask, rows, cols };
  }
  function pickByMask(rng, m) {
    const mat = toMatrix(rng);
    if (mat.length !== m.rows || mat[0].length !== m.cols) {
      // Excel: zakres sumy o innym rozmiarze — bierze ten sam kształt od lewego górnego rogu
      if (isRange(rng)) return pickByMask(new Range(rng.ctx, rng.sheet, rng.r1, rng.c1, rng.r1 + m.rows - 1, rng.c1 + m.cols - 1), m);
      return ERR["#VALUE!"];
    }
    const out = [];
    for (let i = 0; i < m.rows; i++) for (let j = 0; j < m.cols; j++) if (m.mask[i * m.cols + j]) out.push(mat[i][j]);
    return out;
  }
  // Skrót przy przeliczaniu wielu formuł (ctx.matrixCache): LICZ.JEŻELI / SUMA.WARUNKÓW
  // z JEDNYM warunkiem „równa się tekst” po tym samym zakresie (np. 148 × LICZ.JEŻELI(tabela;
  // C5) w arkuszu analiz) — zamiast przechodzić zakres w każdej formule, raz liczymy
  // „tekst → ile / suma”. Tylko gdy semantyka jest identyczna z makeCriterion: tekst bez
  // operatora (albo z „=”), bez symboli wieloznacznych, nie liczba i nie PRAWDA/FAŁSZ —
  // wtedy pasują wyłącznie komórki tekstowe równe bez względu na wielkość liter.
  function textEqKey(crit) {
    if (typeof crit !== "string" || /^(<=|>=|<>|<|>)/.test(crit)) return null;
    const rhs = crit[0] === "=" ? crit.slice(1) : crit;
    if (rhs === "" || /[*?~]/.test(rhs) || parseNumberText(rhs) !== null || /^(TRUE|PRAWDA|FALSE|FAŁSZ)$/i.test(rhs)) return null;
    return rhs.toLowerCase();
  }
  const cachedMatrix = (v) => (isRange(v) && v.ctx.matrixCache ? v.toMatrix() : null);
  const textCountMemo = new WeakMap(); // macierz z pamięci zakresów → Map(tekst → ile)
  function textCounts(mat) {
    let m = textCountMemo.get(mat);
    if (!m) {
      m = new Map();
      for (const row of mat) for (const v of row) if (typeof v === "string") { const k = v.toLowerCase(); m.set(k, (m.get(k) || 0) + 1); }
      textCountMemo.set(mat, m);
    }
    return m;
  }
  const textSumMemo = new WeakMap(); // macierz warunku → WeakMap(macierz sumy → Map(tekst → {sum, err}))
  function textSums(critMat, sumMat) {
    let bySum = textSumMemo.get(critMat);
    if (!bySum) { bySum = new WeakMap(); textSumMemo.set(critMat, bySum); }
    let m = bySum.get(sumMat);
    if (!m) {
      m = new Map();
      for (let i = 0; i < critMat.length; i++) for (let j = 0; j < critMat[i].length; j++) {
        const c = critMat[i][j];
        if (typeof c !== "string") continue;
        const k = c.toLowerCase();
        let g = m.get(k);
        if (!g) { g = { sum: 0, err: null }; m.set(k, g); }
        const v = sumMat[i][j];
        if (isErr(v)) { if (!g.err) g.err = v; } else if (typeof v === "number") g.sum += v;
      }
      bySum.set(sumMat, m);
    }
    return m;
  }
  FN.COUNTIF = (args) => {
    const mat = cachedMatrix(args[0]);
    const key = mat ? textEqKey(scalar(args[1])) : null;
    if (key !== null) return textCounts(mat).get(key) || 0;
    const m = criteriaMask([[args[0], args[1]]]);
    return isErr(m) ? m : m.mask.reduce((a, b) => a + b, 0);
  };
  FN.COUNTIFS = (args) => {
    const pairs = [];
    for (let i = 0; i < args.length; i += 2) pairs.push([args[i], args[i + 1]]);
    const m = criteriaMask(pairs);
    return isErr(m) ? m : m.mask.reduce((a, b) => a + b, 0);
  };
  const sumNums = (vals) => vals.reduce((a, v) => a + (typeof v === "number" ? v : 0), 0);
  FN.SUMIF = (args) => {
    const m = criteriaMask([[args[0], args[1]]]);
    if (isErr(m)) return m;
    const vals = pickByMask(args[2] === undefined || (args[2] && args[2].t === "missing") ? args[0] : args[2], m);
    if (isErr(vals)) return vals;
    const e = vals.find(isErr);
    return e || sumNums(vals);
  };
  FN.SUMIFS = (args) => {
    if (args.length === 3) {
      const critMat = cachedMatrix(args[1]);
      const sumMat = critMat ? cachedMatrix(args[0]) : null;
      const key = sumMat && args[0].rows === args[1].rows && args[0].cols === args[1].cols ? textEqKey(scalar(args[2])) : null;
      if (key !== null) { const g = textSums(critMat, sumMat).get(key); return g ? (g.err || g.sum) : 0; }
    }
    const pairs = [];
    for (let i = 1; i < args.length; i += 2) pairs.push([args[i], args[i + 1]]);
    const m = criteriaMask(pairs);
    if (isErr(m)) return m;
    const vals = pickByMask(args[0], m);
    if (isErr(vals)) return vals;
    const e = vals.find(isErr);
    return e || sumNums(vals);
  };
  const avgOf = (vals) => {
    const e = vals.find(isErr);
    if (e) return e;
    const ns = vals.filter((v) => typeof v === "number");
    return ns.length ? ns.reduce((a, b) => a + b, 0) / ns.length : ERR["#DIV/0!"];
  };
  FN.AVERAGEIF = (args) => {
    const m = criteriaMask([[args[0], args[1]]]);
    if (isErr(m)) return m;
    const vals = pickByMask(args[2] === undefined || (args[2] && args[2].t === "missing") ? args[0] : args[2], m);
    return isErr(vals) ? vals : avgOf(vals);
  };
  FN.AVERAGEIFS = (args) => {
    const pairs = [];
    for (let i = 1; i < args.length; i += 2) pairs.push([args[i], args[i + 1]]);
    const m = criteriaMask(pairs);
    if (isErr(m)) return m;
    const vals = pickByMask(args[0], m);
    return isErr(vals) ? vals : avgOf(vals);
  };
  const ifsExtreme = (pick) => (args) => {
    const pairs = [];
    for (let i = 1; i < args.length; i += 2) pairs.push([args[i], args[i + 1]]);
    const m = criteriaMask(pairs);
    if (isErr(m)) return m;
    const vals = pickByMask(args[0], m);
    if (isErr(vals)) return vals;
    const ns = vals.filter((v) => typeof v === "number");
    return ns.length ? pick(ns) : 0;
  };
  FN.MAXIFS = ifsExtreme(maxOf);
  FN.MINIFS = ifsExtreme(minOf);

  FN.SUMPRODUCT = (args) => {
    const mats = args.map(toMatrix);
    const rows = mats[0].length, cols = mats[0][0].length;
    if (mats.some((m) => m.length !== rows || m[0].length !== cols)) return ERR["#VALUE!"];
    let s = 0;
    for (let i = 0; i < rows; i++) for (let j = 0; j < cols; j++) {
      let p = 1;
      for (const m of mats) {
        const v = m[i][j];
        if (isErr(v)) return v;
        p *= typeof v === "number" ? v : (typeof v === "boolean" ? 0 : 0);
      }
      s += p;
    }
    return s;
  };

  const n1 = (f) => (args, ctx) => {
    const x = num(scalar(args[0], ctx));
    if (isErr(x)) return x;
    const r = f(x);
    return typeof r === "number" && !Number.isFinite(r) ? ERR["#NUM!"] : r;
  };
  const optNum = (a, ctx, d) => (a === undefined || (a && a.t === "missing") ? d : num(scalar(a, ctx)));
  function roundHalfAway(x, digits) { return roundExcel(x, digits, "half"); }
  const roundFn = (mode) => (args, ctx) => {
    const x = num(scalar(args[0], ctx)), d = optNum(args[1], ctx, 0);
    if (isErr(x)) return x; if (isErr(d)) return d;
    return roundExcel(x, d, mode);
  };
  FN.ROUND = roundFn("half");
  FN.ROUNDUP = roundFn("up");
  FN.ROUNDDOWN = roundFn("down");
  FN.INT = n1(Math.floor);
  FN.TRUNC = (args, ctx) => FN.ROUNDDOWN([args[0], args[1] === undefined ? 0 : args[1]], ctx);
  FN.ABS = n1(Math.abs);
  FN.SIGN = n1(Math.sign);
  FN.SQRT = n1((x) => (x < 0 ? ERR["#NUM!"] : Math.sqrt(x)));
  FN.EXP = n1(Math.exp);
  FN.LN = n1((x) => (x <= 0 ? ERR["#NUM!"] : Math.log(x)));
  FN.LOG10 = n1((x) => (x <= 0 ? ERR["#NUM!"] : Math.log10(x)));
  FN.LOG = (args, ctx) => {
    const x = num(scalar(args[0], ctx)), b = optNum(args[1], ctx, 10);
    if (isErr(x)) return x; if (isErr(b)) return b;
    if (x <= 0 || b <= 0 || b === 1) return ERR["#NUM!"];
    return Math.log(x) / Math.log(b);
  };
  FN.PI = () => Math.PI;
  FN.POWER = (args, ctx) => {
    const a = num(scalar(args[0], ctx)), b = num(scalar(args[1], ctx));
    if (isErr(a)) return a; if (isErr(b)) return b;
    return excelPow(a, b);
  };
  FN.MOD = (args, ctx) => {
    const a = num(scalar(args[0], ctx)), b = num(scalar(args[1], ctx));
    if (isErr(a)) return a; if (isErr(b)) return b;
    if (b === 0) return ERR["#DIV/0!"];
    return a - b * Math.floor(a / b);
  };
  const toMultiple = (fn) => (args, ctx) => {
    const x = num(scalar(args[0], ctx)), s = optNum(args[1], ctx, x < 0 ? -1 : 1);
    if (isErr(x)) return x; if (isErr(s)) return s;
    if (s === 0) return 0;
    if (x > 0 && s < 0) return ERR["#NUM!"];
    const q = x / s;
    const near = Math.round(q);
    return (Math.abs(q - near) < 1e-9 ? near : fn(q)) * s;
  };
  FN.CEILING = toMultiple(Math.ceil);
  FN.FLOOR = toMultiple(Math.floor);
  FN["CEILING.MATH"] = (args, ctx) => {
    const x = num(scalar(args[0], ctx)), s = Math.abs(optNum(args[1], ctx, 1));
    if (isErr(x)) return x;
    if (s === 0) return 0;
    return Math.ceil(x / s - 1e-12) * s;
  };
  FN["FLOOR.MATH"] = (args, ctx) => {
    const x = num(scalar(args[0], ctx)), s = Math.abs(optNum(args[1], ctx, 1));
    if (isErr(x)) return x;
    if (s === 0) return 0;
    return Math.floor(x / s + 1e-12) * s;
  };
  FN.MROUND = (args, ctx) => {
    const x = num(scalar(args[0], ctx)), m = num(scalar(args[1], ctx));
    if (isErr(x)) return x; if (isErr(m)) return m;
    if (m === 0) return 0;
    if (x * m < 0) return ERR["#NUM!"];
    return roundHalfAway(x / m, 0) * m;
  };
  FN.EVEN = n1((x) => { const c = Math.ceil(Math.abs(x) / 2) * 2; return x < 0 ? -c : c; });
  FN.ODD = n1((x) => { let c = Math.ceil(Math.abs(x)); if (c % 2 === 0) c += 1; return x < 0 ? -c : c; });
  FN.RAND = () => Math.random();
  FN.RANDBETWEEN = (args, ctx) => {
    const a = Math.ceil(num(scalar(args[0], ctx))), b = Math.floor(num(scalar(args[1], ctx)));
    if (!Number.isFinite(a) || !Number.isFinite(b) || a > b) return ERR["#NUM!"];
    return a + Math.floor(Math.random() * (b - a + 1));
  };
  const kth = (desc) => (args, ctx) => {
    const ns = numbersFrom([args[0]]);
    if (isErr(ns)) return ns;
    const k = num(scalar(args[1], ctx));
    if (isErr(k)) return k;
    const kk = Math.ceil(k);
    if (kk < 1 || kk > ns.length) return ERR["#NUM!"];
    const s = ns.slice().sort((a, b) => (desc ? b - a : a - b));
    return s[kk - 1];
  };
  FN.LARGE = kth(true);
  FN.SMALL = kth(false);
  FN.RANK = FN["RANK.EQ"] = (args, ctx) => {
    const x = num(scalar(args[0], ctx));
    if (isErr(x)) return x;
    const ns = numbersFrom([args[1]]);
    if (isErr(ns)) return ns;
    const asc = !!(args[2] !== undefined && !(args[2] && args[2].t === "missing") && num(scalar(args[2], ctx)));
    if (!ns.includes(x)) return ERR["#N/A"];
    return 1 + ns.filter((v) => (asc ? v < x : v > x)).length;
  };
  FN.SUBTOTAL = (args, ctx) => {
    const code = num(scalar(args[0], ctx));
    if (isErr(code)) return code;
    const map = { 1: "AVERAGE", 2: "COUNT", 3: "COUNTA", 4: "MAX", 5: "MIN", 6: "PRODUCT", 7: "STDEV", 8: "STDEV.P", 9: "SUM", 10: "VAR", 11: "VAR.P" };
    const fn = map[code > 100 ? code - 100 : code];
    if (!fn) return ERR["#VALUE!"];
    // Jak Excel: komórki z własnym SUMY.CZĘŚCIOWE w zakresie są pomijane — inaczej suma
    // końcowa nad grupami z sumami częściowymi liczyła wszystko podwójnie (i trafiało to do <v>).
    const isSub = typeof ctx.formula === "function"
      ? (sh, r, c) => /\b(?:_xlfn\.)?(?:SUBTOTAL|AGGREGATE)\s*\(/i.test(ctx.formula(sh, r, c) || "")
      : null;
    const parts = args.slice(1).map((a) => {
      if (!isSub || !isRange(a)) return a;
      const fr = new Range(a.ctx, a.sheet, a.r1, a.c1, a.r2, a.c2);
      fr.at = (i, j) => (isSub(a.sheet, a.r1 + i, a.c1 + j) ? null : Range.prototype.at.call(fr, i, j));
      return fr;
    });
    return FN[fn](parts, ctx);
  };

  // logiczne (leniwe)
  const isMissing = (n) => !n || n.t === "missing";
  LAZY.IF = (nodes, ctx) => {
    const c = scalar(evalNode(nodes[0], ctx), ctx);
    if (isMatrix(c) || isRange(c)) return ERR["#VALUE!"];
    const b = bool(c);
    if (isErr(b)) return b;
    if (b) return isMissing(nodes[1]) ? (nodes.length > 1 ? 0 : true) : evalNode(nodes[1], ctx);
    if (nodes.length < 3) return false;
    return isMissing(nodes[2]) ? 0 : evalNode(nodes[2], ctx);
  };
  LAZY.IFS = (nodes, ctx) => {
    for (let i = 0; i + 1 < nodes.length; i += 2) {
      const b = bool(scalar(evalNode(nodes[i], ctx), ctx));
      if (isErr(b)) return b;
      if (b) return evalNode(nodes[i + 1], ctx);
    }
    return ERR["#N/A"];
  };
  LAZY.IFERROR = (nodes, ctx) => {
    const v = evalNode(nodes[0], ctx);
    const s = isRange(v) ? scalar(v, ctx) : v;
    return isErr(s) ? evalNode(nodes[1], ctx) : v;
  };
  LAZY.IFNA = (nodes, ctx) => {
    const v = evalNode(nodes[0], ctx);
    const s = isRange(v) ? scalar(v, ctx) : v;
    return isErr(s) && s.err === "#N/A" ? evalNode(nodes[1], ctx) : v;
  };
  LAZY.SWITCH = (nodes, ctx) => {
    const x = scalar(evalNode(nodes[0], ctx), ctx);
    if (isErr(x)) return x;
    let i = 1;
    for (; i + 1 < nodes.length; i += 2) {
      const c = scalar(evalNode(nodes[i], ctx), ctx);
      if (!isErr(c) && compare(x, c) === 0 && typeRank(x) === typeRank(c)) return evalNode(nodes[i + 1], ctx);
    }
    return i < nodes.length ? evalNode(nodes[i], ctx) : ERR["#N/A"];
  };
  LAZY.CHOOSE = (nodes, ctx) => {
    const k = num(scalar(evalNode(nodes[0], ctx), ctx));
    if (isErr(k)) return k;
    const i = Math.trunc(k);
    if (i < 1 || i >= nodes.length) return ERR["#VALUE!"];
    return evalNode(nodes[i], ctx);
  };
  const logicalValues = (args) => {
    const out = [];
    for (const a of args) {
      if (isRange(a) || isMatrix(a)) {
        for (const v of flat(a)) {
          if (isErr(v)) return v;
          if (typeof v === "boolean") out.push(v);
          else if (typeof v === "number") out.push(v !== 0);
        }
      } else if (!(a && a.t === "missing")) {
        const b = bool(a);
        if (isErr(b)) return b;
        out.push(b);
      }
    }
    return out.length ? out : ERR["#VALUE!"];
  };
  FN.AND = (args) => { const v = logicalValues(args); return isErr(v) ? v : v.every(Boolean); };
  FN.OR = (args) => { const v = logicalValues(args); return isErr(v) ? v : v.some(Boolean); };
  FN.XOR = (args) => { const v = logicalValues(args); return isErr(v) ? v : v.filter(Boolean).length % 2 === 1; };
  FN.NOT = (args, ctx) => { const b = bool(scalar(args[0], ctx)); return isErr(b) ? b : !b; };
  FN.TRUE = () => true;
  FN.FALSE = () => false;

  // informacyjne
  const info = (f) => (args, ctx) => f(scalar(args[0], ctx));
  FN.ISBLANK = info((v) => v === null || v === undefined);
  FN.ISNUMBER = info((v) => typeof v === "number");
  FN.ISTEXT = info((v) => typeof v === "string");
  FN.ISNONTEXT = info((v) => typeof v !== "string");
  FN.ISLOGICAL = info((v) => typeof v === "boolean");
  FN.ISERROR = info(isErr);
  FN.ISERR = info((v) => isErr(v) && v.err !== "#N/A");
  FN.ISNA = info((v) => isErr(v) && v.err === "#N/A");
  FN.ISEVEN = (args, ctx) => { const x = num(scalar(args[0], ctx)); return isErr(x) ? x : Math.trunc(x) % 2 === 0; };
  FN.ISODD = (args, ctx) => { const x = num(scalar(args[0], ctx)); return isErr(x) ? x : Math.abs(Math.trunc(x) % 2) === 1; };
  FN.NA = () => ERR["#N/A"];
  FN.N = info((v) => (typeof v === "number" ? v : typeof v === "boolean" ? (v ? 1 : 0) : isErr(v) ? v : 0));
  FN.T = info((v) => (typeof v === "string" ? v : isErr(v) ? v : ""));

  // tekst
  const s1 = (f) => (args, ctx) => {
    const v = scalar(args[0], ctx);
    if (isErr(v)) return v;
    return f(text(v, ctx), args, ctx);
  };
  FN.LEN = s1((s) => [...s].length);
  FN.UPPER = s1((s) => s.toUpperCase());
  FN.LOWER = s1((s) => s.toLowerCase());
  FN.PROPER = s1((s) => s.toLowerCase().replace(/(^|[^\p{L}])(\p{L})/gu, (m, a, b) => a + b.toUpperCase()));
  FN.TRIM = s1((s) => s.replace(/ +/g, " ").replace(/^ | $/g, ""));
  FN.CLEAN = s1((s) => s.replace(/[\x00-\x1F]/g, ""));
  FN.LEFT = (args, ctx) => {
    const v = scalar(args[0], ctx); if (isErr(v)) return v;
    const n = optNum(args[1], ctx, 1); if (isErr(n)) return n;
    if (n < 0) return ERR["#VALUE!"];
    return [...text(v, ctx)].slice(0, Math.trunc(n)).join("");
  };
  FN.RIGHT = (args, ctx) => {
    const v = scalar(args[0], ctx); if (isErr(v)) return v;
    const n = optNum(args[1], ctx, 1); if (isErr(n)) return n;
    if (n < 0) return ERR["#VALUE!"];
    const ch = [...text(v, ctx)];
    const k = Math.trunc(n);
    return k === 0 ? "" : ch.slice(Math.max(0, ch.length - k)).join("");
  };
  FN.MID = (args, ctx) => {
    const v = scalar(args[0], ctx); if (isErr(v)) return v;
    const st = num(scalar(args[1], ctx)), n = num(scalar(args[2], ctx));
    if (isErr(st)) return st; if (isErr(n)) return n;
    if (st < 1 || n < 0) return ERR["#VALUE!"];
    return [...text(v, ctx)].slice(Math.trunc(st) - 1, Math.trunc(st) - 1 + Math.trunc(n)).join("");
  };
  FN.CONCATENATE = (args, ctx) => {
    let out = "";
    for (const a of args) { const v = scalar(a, ctx); if (isErr(v)) return v; out += text(v, ctx); }
    return out;
  };
  FN.CONCAT = (args, ctx) => {
    let out = "";
    for (const a of args) for (const v of flat(a)) { if (isErr(v)) return v; out += text(v, ctx); }
    return out;
  };
  FN.TEXTJOIN = (args, ctx) => {
    const sep = text(scalar(args[0], ctx), ctx);
    const ignore = bool(scalar(args[1], ctx));
    const parts = [];
    for (const a of args.slice(2)) for (const v of flat(a)) {
      if (isErr(v)) return v;
      const t = text(v, ctx);
      if (ignore && t === "") continue;
      parts.push(t);
    }
    return parts.join(sep);
  };
  FN.REPT = (args, ctx) => {
    const v = scalar(args[0], ctx); if (isErr(v)) return v;
    const n = num(scalar(args[1], ctx)); if (isErr(n)) return n;
    if (n < 0) return ERR["#VALUE!"];
    return text(v, ctx).repeat(Math.trunc(n));
  };
  FN.EXACT = (args, ctx) => text(scalar(args[0], ctx), ctx) === text(scalar(args[1], ctx), ctx);
  const findLike = (ci) => (args, ctx) => {
    const needle = scalar(args[0], ctx), hay = scalar(args[1], ctx);
    if (isErr(needle)) return needle; if (isErr(hay)) return hay;
    const st = optNum(args[2], ctx, 1);
    if (isErr(st)) return st;
    const h = text(hay, ctx), n = text(needle, ctx);
    if (st < 1 || st > h.length + 1) return ERR["#VALUE!"];
    let idx;
    if (ci) {
      if (/[*?~]/.test(n)) {
        const re = new RegExp(wildcardRe(n).source.slice(1, -1), "i");
        const m = re.exec(h.slice(st - 1));
        idx = m ? m.index + st - 1 : -1;
      } else idx = h.toLowerCase().indexOf(n.toLowerCase(), st - 1);
    } else idx = h.indexOf(n, st - 1);
    return idx < 0 ? ERR["#VALUE!"] : idx + 1;
  };
  FN.FIND = findLike(false);
  FN.SEARCH = findLike(true);
  FN.SUBSTITUTE = (args, ctx) => {
    const v = scalar(args[0], ctx); if (isErr(v)) return v;
    const s = text(v, ctx), from = text(scalar(args[1], ctx), ctx), to = text(scalar(args[2], ctx), ctx);
    if (!from) return s;
    if (args[3] === undefined || (args[3] && args[3].t === "missing")) return s.split(from).join(to);
    const k = num(scalar(args[3], ctx));
    if (isErr(k)) return k;
    if (k < 1) return ERR["#VALUE!"];
    let idx = -1;
    for (let i = 0; i < k; i++) { idx = s.indexOf(from, idx + 1); if (idx < 0) return s; }
    return s.slice(0, idx) + to + s.slice(idx + from.length);
  };
  FN.REPLACE = (args, ctx) => {
    const s = text(scalar(args[0], ctx), ctx);
    const st = num(scalar(args[1], ctx)), n = num(scalar(args[2], ctx));
    if (isErr(st)) return st; if (isErr(n)) return n;
    if (st < 1 || n < 0) return ERR["#VALUE!"];
    return s.slice(0, st - 1) + text(scalar(args[3], ctx), ctx) + s.slice(st - 1 + n);
  };
  FN.CHAR = n1((x) => (x < 1 || x > 255 ? ERR["#VALUE!"] : String.fromCharCode(x)));
  FN.CODE = s1((s) => (s.length ? s.charCodeAt(0) : ERR["#VALUE!"]));
  FN.VALUE = (args, ctx) => {
    const v = scalar(args[0], ctx);
    if (typeof v === "number") return v;
    if (isErr(v)) return v;
    const n = parseNumberText(text(v, ctx));
    return n === null ? ERR["#VALUE!"] : n;
  };
  FN.TEXT = (args, ctx) => {
    const v = scalar(args[0], ctx);
    if (isErr(v)) return v;
    let fmt = text(scalar(args[1], ctx), ctx);
    // Excel PL zapisuje kody formatu po polsku (rrrr = rok, gg = godziny)
    // — ale nie w tekście w cudzysłowie, w [Red]/[$-415] ani po „\\” („0 ""hrs""” zostaje, [Red] nie robi się [Yed])
    fmt = fmt.replace(/"[^"]*"|\[[^\]]*\]|\\.|[rg]+/gi, (m, at, all) => {
      if (!/^[rg]/i.test(m)) return m;
      return m.replace(/r/gi, (c) => (c === "R" ? "Y" : "y"))
        .replace(/g+$/i, (run) => (/^(?::|$)/.test(all.slice(at + m.length, at + m.length + 1)) ? run.replace(/g/gi, "h") : run.slice(0, -1).replace(/g/gi, "h") + run.slice(-1)));
    });
    const x = typeof v === "number" ? v : parseNumberText(String(v ?? ""));
    if (x === null) return text(v, ctx);
    if (ctx.formatNumber) {
      const out = ctx.formatNumber(fmt, x);
      if (out != null) return out;
    }
    return generalText(x);
  };

  // daty
  FN.TODAY = (args, ctx) => Math.floor(ctx.now());
  FN.NOW = (args, ctx) => ctx.now();
  FN.DATE = (args, ctx) => {
    let y = num(scalar(args[0], ctx)), m = num(scalar(args[1], ctx)), d = num(scalar(args[2], ctx));
    if (isErr(y)) return y; if (isErr(m)) return m; if (isErr(d)) return d;
    y = Math.trunc(y); m = Math.trunc(m); d = Math.trunc(d);
    if (y < 1900) y += 1900;
    const s = (Date.UTC(y, m - 1, d) - EPOCH) / DAY_MS;
    return s < 0 ? ERR["#NUM!"] : s;
  };
  const datePart = (f) => (args, ctx) => {
    const v = scalar(args[0], ctx);
    const x = num(v);
    if (isErr(x)) return x;
    if (x < 0) return ERR["#NUM!"];
    return f(serialToParts(x), x);
  };
  FN.YEAR = datePart((p) => p.y);
  FN.MONTH = datePart((p) => p.m);
  FN.DAY = datePart((p) => p.d);
  FN.HOUR = datePart((p, x) => Math.floor(((x % 1) * 24) + 1e-9));
  FN.MINUTE = datePart((p, x) => Math.floor((((x % 1) * 1440) + 1e-9) % 60));
  FN.SECOND = datePart((p, x) => Math.round(((x % 1) * 86400)) % 60);
  FN.TIME = (args, ctx) => {
    const h = num(scalar(args[0], ctx)), m = num(scalar(args[1], ctx)), s = num(scalar(args[2], ctx));
    if (isErr(h)) return h; if (isErr(m)) return m; if (isErr(s)) return s;
    const t = (h * 3600 + m * 60 + s) / 86400;
    return t - Math.floor(t);
  };
  FN.WEEKDAY = (args, ctx) => {
    const x = num(scalar(args[0], ctx)); if (isErr(x)) return x;
    const type = optNum(args[1], ctx, 1);
    const wd = serialToParts(x).wd; // 0 = niedziela
    if (type === 1 || type === 17) return wd + 1;
    if (type === 2 || type === 11) return wd === 0 ? 7 : wd;
    if (type === 3) return wd === 0 ? 6 : wd - 1;
    if (type >= 12 && type <= 16) { const start = type - 10; return ((wd - start + 7) % 7) + 1; }
    return ERR["#NUM!"];
  };
  FN.WEEKNUM = (args, ctx) => {
    const x = num(scalar(args[0], ctx)); if (isErr(x)) return x;
    const type = optNum(args[1], ctx, 1);
    if (type === 21) return FN.ISOWEEKNUM([x], ctx);
    const p = serialToParts(x);
    const jan1 = dateSerial(p.y, 1, 1);
    const startDay = type === 2 || type === 11 ? 1 : type >= 12 && type <= 17 ? type - 10 : 0;
    const wdJan1 = serialToParts(jan1).wd;
    const offset = (wdJan1 - startDay + 7) % 7;
    return Math.floor((Math.floor(x) - jan1 + offset) / 7) + 1;
  };
  FN.ISOWEEKNUM = (args, ctx) => {
    const x = num(scalar(args[0], ctx)); if (isErr(x)) return x;
    const d = new Date(EPOCH + Math.floor(x) * DAY_MS);
    const day = (d.getUTCDay() + 6) % 7;
    d.setUTCDate(d.getUTCDate() - day + 3);
    const firstThu = new Date(Date.UTC(d.getUTCFullYear(), 0, 4));
    return 1 + Math.round(((d - firstThu) / DAY_MS - 3 + ((firstThu.getUTCDay() + 6) % 7)) / 7);
  };
  FN.EDATE = (args, ctx) => {
    const x = num(scalar(args[0], ctx)), k = num(scalar(args[1], ctx));
    if (isErr(x)) return x; if (isErr(k)) return k;
    const p = serialToParts(x);
    const m = p.m - 1 + Math.trunc(k);
    const y = p.y + Math.floor(m / 12);
    const mm = ((m % 12) + 12) % 12 + 1;
    return dateSerial(y, mm, Math.min(p.d, daysInMonth(y, mm)));
  };
  FN.EOMONTH = (args, ctx) => {
    const x = num(scalar(args[0], ctx)), k = num(scalar(args[1], ctx));
    if (isErr(x)) return x; if (isErr(k)) return k;
    const p = serialToParts(x);
    const m = p.m - 1 + Math.trunc(k);
    const y = p.y + Math.floor(m / 12);
    const mm = ((m % 12) + 12) % 12 + 1;
    return dateSerial(y, mm, daysInMonth(y, mm));
  };
  FN.DAYS = (args, ctx) => {
    const a = num(scalar(args[0], ctx)), b = num(scalar(args[1], ctx));
    if (isErr(a)) return a; if (isErr(b)) return b;
    return Math.floor(a) - Math.floor(b);
  };
  FN.DATEDIF = (args, ctx) => {
    const a = num(scalar(args[0], ctx)), b = num(scalar(args[1], ctx));
    if (isErr(a)) return a; if (isErr(b)) return b;
    const unit = text(scalar(args[2], ctx), ctx).toUpperCase();
    const s = Math.floor(a), e = Math.floor(b);
    if (s > e) return ERR["#NUM!"];
    const ps = serialToParts(s), pe = serialToParts(e);
    let months = (pe.y - ps.y) * 12 + (pe.m - ps.m);
    if (pe.d < ps.d) months -= 1;
    switch (unit) {
      case "D": return e - s;
      case "M": return months;
      case "Y": return Math.floor(months / 12);
      case "YM": return months % 12;
      case "MD": {
        if (pe.d >= ps.d) return pe.d - ps.d;
        const pm = pe.m === 1 ? 12 : pe.m - 1, py = pe.m === 1 ? pe.y - 1 : pe.y;
        return daysInMonth(py, pm) - ps.d + pe.d;
      }
      case "YD": {
        let st = dateSerial(pe.y, ps.m, Math.min(ps.d, daysInMonth(pe.y, ps.m)));
        if (st > e) st = dateSerial(pe.y - 1, ps.m, Math.min(ps.d, daysInMonth(pe.y - 1, ps.m)));
        return e - st;
      }
      default: return ERR["#NUM!"];
    }
  };
  const holidaySet = (arg) => {
    const out = new Set();
    if (arg === undefined || (arg && arg.t === "missing")) return out;
    for (const v of flat(arg)) if (typeof v === "number") out.add(Math.floor(v));
    return out;
  };
  const isWorkday = (s, hol) => { const wd = serialToParts(s).wd; return wd !== 0 && wd !== 6 && !hol.has(s); };
  FN.NETWORKDAYS = (args, ctx) => {
    const a = num(scalar(args[0], ctx)), b = num(scalar(args[1], ctx));
    if (isErr(a)) return a; if (isErr(b)) return b;
    const hol = holidaySet(args[2]);
    let s = Math.floor(a), e = Math.floor(b), sign = 1;
    if (s > e) { [s, e] = [e, s]; sign = -1; }
    let n = 0;
    for (let d = s; d <= e; d++) if (isWorkday(d, hol)) n++;
    return n * sign;
  };
  FN.WORKDAY = (args, ctx) => {
    const a = num(scalar(args[0], ctx)), k = num(scalar(args[1], ctx));
    if (isErr(a)) return a; if (isErr(k)) return k;
    const hol = holidaySet(args[2]);
    let d = Math.floor(a), left = Math.trunc(k);
    const step = left < 0 ? -1 : 1;
    while (left !== 0) { d += step; if (isWorkday(d, hol)) left -= step; }
    return d;
  };
  FN.DATEVALUE = (args, ctx) => {
    const v = scalar(args[0], ctx);
    if (isErr(v)) return v;
    const n = parseNumberText(text(v, ctx));
    return n === null ? ERR["#VALUE!"] : Math.floor(n);
  };

  // wyszukiwanie
  function lookupMatch(vals, x, mode) {
    // mode 0 = dokładne (z * ? dla tekstu), 1 = największe ≤ (posortowane rosnąco), -1 = najmniejsze ≥ (malejąco)
    if (mode === 0) {
      if (typeof x === "string" && /[*?~]/.test(x)) {
        const re = wildcardRe(x);
        return vals.findIndex((v) => typeof v === "string" && re.test(v));
      }
      return vals.findIndex((v) => v !== null && typeRank(v) === typeRank(x) && compare(v, x) === 0);
    }
    let found = -1;
    for (let i = 0; i < vals.length; i++) {
      const v = vals[i];
      if (v === null || typeRank(v) !== typeRank(x)) continue;
      const c = compare(v, x);
      if (mode === 1) { if (c <= 0) found = i; else break; }
      else { if (c >= 0) found = i; else break; }
    }
    return found;
  }
  FN.VLOOKUP = (args, ctx) => {
    const x = scalar(args[0], ctx);
    if (isErr(x)) return x;
    const tbl = toMatrix(args[1]);
    const col = num(scalar(args[2], ctx));
    if (isErr(col)) return col;
    if (col < 1) return ERR["#VALUE!"];
    if (col > tbl[0].length) return ERR["#REF!"];
    const approx = args[3] === undefined || (args[3] && args[3].t === "missing") ? true : bool(scalar(args[3], ctx));
    const i = lookupMatch(tbl.map((r) => r[0]), x, approx ? 1 : 0);
    if (i < 0) return ERR["#N/A"];
    const v = tbl[i][Math.trunc(col) - 1];
    return v === null ? 0 : v;
  };
  FN.HLOOKUP = (args, ctx) => {
    const x = scalar(args[0], ctx);
    if (isErr(x)) return x;
    const tbl = toMatrix(args[1]);
    const row = num(scalar(args[2], ctx));
    if (isErr(row)) return row;
    if (row < 1) return ERR["#VALUE!"];
    if (row > tbl.length) return ERR["#REF!"];
    const approx = args[3] === undefined || (args[3] && args[3].t === "missing") ? true : bool(scalar(args[3], ctx));
    const i = lookupMatch(tbl[0], x, approx ? 1 : 0);
    if (i < 0) return ERR["#N/A"];
    const v = tbl[Math.trunc(row) - 1][i];
    return v === null ? 0 : v;
  };
  const vectorOf = (a) => { const m = toMatrix(a); return m.length === 1 ? m[0] : m.map((r) => r[0]); };
  FN.MATCH = (args, ctx) => {
    const x = scalar(args[0], ctx);
    if (isErr(x)) return x;
    const mode = args[2] === undefined || (args[2] && args[2].t === "missing") ? 1 : Math.sign(num(scalar(args[2], ctx)));
    const i = lookupMatch(vectorOf(args[1]), x, mode);
    return i < 0 ? ERR["#N/A"] : i + 1;
  };
  FN.XMATCH = (args, ctx) => {
    const x = scalar(args[0], ctx);
    if (isErr(x)) return x;
    const vals = vectorOf(args[1]);
    const mode = optNum(args[2], ctx, 0);
    const rev = optNum(args[3], ctx, 1) < 0;
    const order = rev ? vals.map((_, i) => vals.length - 1 - i) : vals.map((_, i) => i);
    let best = -1;
    for (const i of order) {
      const v = vals[i];
      if (mode === 2) {
        if (lookupMatch([v], x, 0) === 0) { best = i; break; }
      } else if (mode === 0) {
        // tryb 0 = dokładnie, BEZ symboli wieloznacznych (to tryb 2) — „a*” szuka tekstu „a*”
        if (v !== null && typeRank(v) === typeRank(x) && compare(v, x) === 0) { best = i; break; }
      } else if (v !== null && typeRank(v) === typeRank(x)) {
        const c = compare(v, x);
        if (c === 0) { best = i; break; }
        if (mode === -1 && c < 0 && (best < 0 || compare(v, vals[best]) > 0)) best = i;
        if (mode === 1 && c > 0 && (best < 0 || compare(v, vals[best]) < 0)) best = i;
      }
    }
    return best < 0 ? ERR["#N/A"] : best + 1;
  };
  FN.XLOOKUP = (args, ctx) => {
    const pos = FN.XMATCH([args[0], args[1], args[4], args[5]], ctx);
    if (isErr(pos)) {
      if (pos.err === "#N/A" && args[3] !== undefined && !(args[3] && args[3].t === "missing")) return args[3];
      return pos;
    }
    const ret = args[2];
    if (isRange(ret)) {
      const vertical = toMatrix(args[1]).length > 1;
      if (vertical) return ret.cols === 1 ? ret.at(pos - 1, 0) : new Range(ret.ctx, ret.sheet, ret.r1 + pos - 1, ret.c1, ret.r1 + pos - 1, ret.c2);
      return ret.rows === 1 ? ret.at(0, pos - 1) : new Range(ret.ctx, ret.sheet, ret.r1, ret.c1 + pos - 1, ret.r2, ret.c1 + pos - 1);
    }
    const v = vectorOf(ret)[pos - 1];
    return v === undefined ? ERR["#N/A"] : v;
  };
  FN.LOOKUP = (args, ctx) => {
    const x = scalar(args[0], ctx);
    if (isErr(x)) return x;
    const look = vectorOf(args[1]);
    const res = args[2] === undefined ? look : vectorOf(args[2]);
    const i = lookupMatch(look, x, 1);
    return i < 0 ? ERR["#N/A"] : (res[i] === undefined ? ERR["#N/A"] : res[i]);
  };
  FN.INDEX = (args, ctx) => {
    const a = args[0];
    let r = optNum(args[1], ctx, 0), c = optNum(args[2], ctx, 0);
    if (isErr(r)) return r; if (isErr(c)) return c;
    r = Math.trunc(r); c = Math.trunc(c);
    if (isRange(a)) {
      if (a.rows === 1 && args[2] === undefined && r > 0) { c = r; r = 1; }
      if (r < 0 || c < 0 || r > a.rows || c > a.cols) return ERR["#REF!"];
      if (r === 0 && c === 0) return a;
      if (r === 0) return new Range(a.ctx, a.sheet, a.r1, a.c1 + c - 1, a.r2, a.c1 + c - 1);
      if (c === 0) return a.cols === 1 ? a.at(r - 1, 0) : new Range(a.ctx, a.sheet, a.r1 + r - 1, a.c1, a.r1 + r - 1, a.c2);
      return a.at(r - 1, c - 1);
    }
    const m = toMatrix(a);
    if (m.length === 1 && args[2] === undefined) { c = r; r = 1; }
    const row = m[(r || 1) - 1];
    if (!row) return ERR["#REF!"];
    const v = row[(c || 1) - 1];
    return v === undefined ? ERR["#REF!"] : v;
  };
  FN.ROWS = (args) => toMatrix(args[0]).length;
  FN.COLUMNS = (args) => toMatrix(args[0])[0].length;
  FN.ROW = (args, ctx) => (args.length && isRange(args[0]) ? args[0].r1 + 1 : ctx.row + 1);
  FN.COLUMN = (args, ctx) => (args.length && isRange(args[0]) ? args[0].c1 + 1 : ctx.col + 1);

  // ── ewaluacja ──────────────────────────────────────────────────────────────
  function arith(op, a, b) {
    const x = num(a);
    if (isErr(x)) return x;
    const y = num(b);
    if (isErr(y)) return y;
    switch (op) {
      case "+": { const r = x + y; return Number.isFinite(r) ? r : ERR["#NUM!"]; }
      case "-": { const r = x - y; return Number.isFinite(r) ? r : ERR["#NUM!"]; }
      case "*": { const r = x * y; return Number.isFinite(r) ? r : ERR["#NUM!"]; }
      case "/": { if (y === 0) return ERR["#DIV/0!"]; const r = x / y; return Number.isFinite(r) ? r : ERR["#NUM!"]; }
      case "^": return excelPow(x, y);
      default: return ERR["#VALUE!"];
    }
  }
  function binScalar(op, a, b, ctx) {
    if (op === "&") {
      if (isErr(a)) return a;
      if (isErr(b)) return b;
      return text(a, ctx) + text(b, ctx);
    }
    if (CMP.has(op)) {
      if (isErr(a)) return a;
      if (isErr(b)) return b;
      const c = compare(a, b);
      switch (op) {
        case "=": return c === 0; case "<>": return c !== 0; case "<": return c < 0;
        case ">": return c > 0; case "<=": return c <= 0; case ">=": return c >= 0;
        default: return false;
      }
    }
    return arith(op, a, b);
  }
  // działanie na tablicach (A1:A3*2, (B1:B5>0)*(C1:C5)) — element po elemencie
  function broadcast(fn, a, b) {
    const A = toMatrix(a), B = toMatrix(b);
    const rows = Math.max(A.length, B.length), cols = Math.max(A[0].length, B[0].length);
    const get = (M, i, j) => {
      const ii = M.length === 1 ? 0 : i, jj = M[0].length === 1 ? 0 : j;
      return M[ii] && M[ii][jj] !== undefined ? M[ii][jj] : ERR["#N/A"];
    };
    const out = [];
    for (let i = 0; i < rows; i++) { const row = []; for (let j = 0; j < cols; j++) row.push(fn(get(A, i, j), get(B, i, j))); out.push(row); }
    return out;
  }

  function resolveSheet(ctx, name) { return name == null ? ctx.sheetName : ctx.sheetFor(name); }

  function evalNode(n, ctx) {
    switch (n.t) {
      case "num": case "str": case "bool": return n.v;
      case "err": return errOf(n.v);
      case "missing": return null;
      case "paren": return evalNode(n.a, ctx);
      case "ref": {
        const sh = resolveSheet(ctx, n.sheet);
        if (sh === null) return ERR["#REF!"];
        return new Range(ctx, sh, n.r, n.c, n.r, n.c);
      }
      case "range": {
        const sh = resolveSheet(ctx, n.sheet);
        if (sh === null) return ERR["#REF!"];
        return new Range(ctx, sh, n.r1, n.c1, n.r2, n.c2);
      }
      case "sref": {
        const r = ctx.structured ? ctx.structured(n.table, n.spec, ctx) : null;
        if (!r) return ERR["#REF!"];
        return new Range(ctx, r.sheet, r.r1, r.c1, r.r2, r.c2);
      }
      case "name": {
        const def = ctx.name ? ctx.name(n.v, n.sheet, ctx) : null;
        if (def === null || def === undefined) return ERR["#NAME?"];
        return def;
      }
      case "array": return n.rows.map((r) => r.map((x) => scalar(evalNode(x, ctx), ctx)));
      case "un": {
        const v = evalNode(n.a, ctx);
        if (isRange(v) && (v.rows > 1 || v.cols > 1)) return broadcast((x) => (n.op === "-" ? arith("*", x, -1) : num(x)), v, [[0]]);
        if (n.op === "+") return scalar(v, ctx) ?? 0; // plus jednoargumentowy nic nie zmienia (+PRAWDA = PRAWDA)
        const x = num(scalar(v, ctx));
        if (isErr(x)) return x;
        return -x;
      }
      case "pct": {
        const x = num(scalar(evalNode(n.a, ctx), ctx));
        return isErr(x) ? x : x / 100;
      }
      case "bin": {
        const a = evalNode(n.a, ctx);
        const b = evalNode(n.b, ctx);
        const multi = (v) => (isRange(v) && (v.rows > 1 || v.cols > 1)) || isMatrix(v);
        if (multi(a) || multi(b)) return broadcast((x, y) => binScalar(n.op, x, y, ctx), a, b);
        return binScalar(n.op, scalar(a, ctx), scalar(b, ctx), ctx);
      }
      case "call": {
        if (LAZY[n.name]) return LAZY[n.name](n.args, ctx);
        const f = FN[n.name];
        if (!f) return ctx.unknownFunction ? ctx.unknownFunction(n.name) : ERR["#NAME?"];
        const args = n.args.map((a) => (a.t === "missing" ? a : evalNode(a, ctx)));
        // błąd w argumencie przechodzi dalej (poza funkcjami, które same sprawdzają błędy)
        if (!INFO_FNS.has(n.name)) for (const a of args) if (isErr(a)) return a;
        if (ctx.depth > 60) return ERR["#NUM!"];
        return f(args, ctx);
      }
      default: return ERR["#VALUE!"];
    }
  }
  const INFO_FNS = new Set(["ISERROR", "ISERR", "ISNA", "ISBLANK", "ISNUMBER", "ISTEXT", "ISNONTEXT", "ISLOGICAL", "N", "T"]);

  // Wynik formuły w komórce: zakres → wartość (niejawne przecięcie), pusta → 0.
  function evaluate(ast, ctx) {
    let v;
    try { v = evalNode(ast, ctx); } catch (e) { if (e && e.unsupported) return UNSUPPORTED; throw e; }
    if (v === UNSUPPORTED) return v;
    v = scalar(v, ctx);
    // korzeń formuły „A−B” / „0+A−B”: wynik bliski zera → 0 (jak Excel; w nawiasach już nie)
    if (ast && ast.t === "bin" && (ast.op === "+" || ast.op === "-") && typeof v === "number" && v !== 0) {
      const left = num(scalar(evalNode(ast.a, ctx), ctx));
      if (typeof left === "number") v = snap(left, v);
    }
    if (isMatrix(v)) v = v[0][0];
    if (v === null || v === undefined) return 0;
    if (typeof v === "number" && !Number.isFinite(v)) return ERR["#NUM!"];
    return v;
  }

  // Odwołania formuły (do grafu zależności): [{ sheet|null, r1,c1,r2,c2 }] + tabele + nazwy.
  function refsOf(ast) {
    const out = [];
    const walk = (n) => {
      if (!n || typeof n !== "object") return;
      switch (n.t) {
        case "ref": out.push({ sheet: n.sheet, r1: n.r, c1: n.c, r2: n.r, c2: n.c }); break;
        case "range": out.push({ sheet: n.sheet, r1: n.r1, c1: n.c1, r2: n.r2, c2: n.c2 }); break;
        case "sref": out.push({ table: n.table, spec: n.spec }); break;
        case "name": out.push({ name: n.v, sheet: n.sheet }); break;
        case "call": n.args.forEach(walk); break;
        case "bin": walk(n.a); walk(n.b); break;
        case "un": case "pct": case "paren": walk(n.a); break;
        case "array": n.rows.forEach((r) => r.forEach(walk)); break;
        default: break;
      }
    };
    walk(ast);
    return out;
  }

  function functionsOf(ast, acc = new Set()) {
    if (!ast || typeof ast !== "object") return acc;
    if (ast.t === "call") { acc.add(ast.name); ast.args.forEach((a) => functionsOf(a, acc)); }
    else if (ast.t === "bin") { functionsOf(ast.a, acc); functionsOf(ast.b, acc); }
    else if (ast.a) functionsOf(ast.a, acc);
    else if (ast.t === "array") ast.rows.forEach((r) => r.forEach((x) => functionsOf(x, acc)));
    return acc;
  }
  const isSupported = (name) => !!(FN[name] || LAZY[name]);
  // Funkcje „datowe” — Excel przy formacie Ogólnym sam ustawia komórce format daty.
  const DATE_RESULT_FNS = new Set(["TODAY", "NOW", "DATE", "EDATE", "EOMONTH", "WORKDAY", "DATEVALUE"]);
  function resultLooksLikeDate(ast) {
    let n = ast;
    while (n && n.t === "paren") n = n.a;
    if (!n) return false;
    if (n.t === "call") return DATE_RESULT_FNS.has(n.name);
    if (n.t === "bin" && (n.op === "+" || n.op === "-")) return resultLooksLikeDate(n.a) && !resultLooksLikeDate(n.b);
    return false;
  }

  // Lista do podpowiedzi i okna „Wstaw funkcję”: nazwa, kategoria, składnia.
  const CATALOG = [
    ["SUM", "math", "liczba1; [liczba2]; …"], ["AVERAGE", "stat", "liczba1; [liczba2]; …"], ["COUNT", "stat", "wartość1; …"],
    ["COUNTA", "stat", "wartość1; …"], ["COUNTBLANK", "stat", "zakres"], ["MIN", "stat", "liczba1; …"], ["MAX", "stat", "liczba1; …"],
    ["MEDIAN", "stat", "liczba1; …"], ["LARGE", "stat", "zakres; k"], ["SMALL", "stat", "zakres; k"], ["RANK.EQ", "stat", "liczba; zakres; [kolejność]"],
    ["COUNTIF", "stat", "zakres; kryterium"], ["COUNTIFS", "stat", "zakres1; kryterium1; …"], ["SUMIF", "math", "zakres; kryterium; [suma_zakres]"],
    ["SUMIFS", "math", "suma_zakres; zakres1; kryterium1; …"], ["AVERAGEIF", "stat", "zakres; kryterium; [średnia_zakres]"],
    ["AVERAGEIFS", "stat", "średnia_zakres; zakres1; kryterium1; …"], ["MAXIFS", "stat", "maks_zakres; zakres1; kryterium1; …"],
    ["MINIFS", "stat", "min_zakres; zakres1; kryterium1; …"], ["SUMPRODUCT", "math", "tablica1; [tablica2]; …"], ["SUBTOTAL", "math", "funkcja; zakres1; …"],
    ["ROUND", "math", "liczba; miejsca"], ["ROUNDUP", "math", "liczba; miejsca"], ["ROUNDDOWN", "math", "liczba; miejsca"], ["INT", "math", "liczba"],
    ["TRUNC", "math", "liczba; [miejsca]"], ["ABS", "math", "liczba"], ["MOD", "math", "liczba; dzielnik"], ["POWER", "math", "liczba; potęga"],
    ["SQRT", "math", "liczba"], ["PRODUCT", "math", "liczba1; …"], ["CEILING.MATH", "math", "liczba; [istotność]"], ["FLOOR.MATH", "math", "liczba; [istotność]"],
    ["MROUND", "math", "liczba; wielokrotność"], ["RAND", "math", ""], ["RANDBETWEEN", "math", "dół; góra"], ["STDEV.S", "stat", "liczba1; …"],
    ["IF", "logic", "warunek; [jeśli_prawda]; [jeśli_fałsz]"], ["IFS", "logic", "warunek1; wartość1; …"], ["IFERROR", "logic", "wartość; wartość_jeśli_błąd"],
    ["IFNA", "logic", "wartość; wartość_jeśli_brak"], ["AND", "logic", "logiczna1; …"], ["OR", "logic", "logiczna1; …"], ["NOT", "logic", "logiczna"],
    ["XOR", "logic", "logiczna1; …"], ["SWITCH", "logic", "wyrażenie; wartość1; wynik1; …; [domyślnie]"], ["CHOOSE", "logic", "nr; wartość1; …"],
    ["VLOOKUP", "lookup", "szukana; tabela; nr_kolumny; [przybliżone]"], ["HLOOKUP", "lookup", "szukana; tabela; nr_wiersza; [przybliżone]"],
    ["XLOOKUP", "lookup", "szukana; tablica_szukana; tablica_wynikowa; [jeśli_brak]; [tryb]; [kierunek]"], ["INDEX", "lookup", "tablica; nr_wiersza; [nr_kolumny]"],
    ["MATCH", "lookup", "szukana; tablica; [typ]"], ["XMATCH", "lookup", "szukana; tablica; [tryb]; [kierunek]"], ["LOOKUP", "lookup", "szukana; wektor; [wynik]"],
    ["ROW", "lookup", "[odwołanie]"], ["COLUMN", "lookup", "[odwołanie]"], ["ROWS", "lookup", "tablica"], ["COLUMNS", "lookup", "tablica"],
    ["TODAY", "date", ""], ["NOW", "date", ""], ["DATE", "date", "rok; miesiąc; dzień"], ["YEAR", "date", "data"], ["MONTH", "date", "data"], ["DAY", "date", "data"],
    ["WEEKDAY", "date", "data; [typ]"], ["WEEKNUM", "date", "data; [typ]"], ["EDATE", "date", "data; miesiące"], ["EOMONTH", "date", "data; miesiące"],
    ["DATEDIF", "date", "od; do; jednostka (\"D\" \"M\" \"Y\")"], ["DAYS", "date", "do; od"], ["NETWORKDAYS", "date", "od; do; [święta]"],
    ["WORKDAY", "date", "start; dni; [święta]"], ["HOUR", "date", "czas"], ["MINUTE", "date", "czas"], ["TIME", "date", "godz; min; sek"],
    ["LEN", "text", "tekst"], ["LEFT", "text", "tekst; [ile]"], ["RIGHT", "text", "tekst; [ile]"], ["MID", "text", "tekst; start; ile"],
    ["UPPER", "text", "tekst"], ["LOWER", "text", "tekst"], ["PROPER", "text", "tekst"], ["TRIM", "text", "tekst"], ["CONCAT", "text", "tekst1; …"],
    ["TEXTJOIN", "text", "separator; pomiń_puste; tekst1; …"], ["TEXT", "text", "wartość; format"], ["VALUE", "text", "tekst"],
    ["FIND", "text", "szukany; tekst; [start]"], ["SEARCH", "text", "szukany; tekst; [start]"], ["SUBSTITUTE", "text", "tekst; stary; nowy; [które]"],
    ["REPLACE", "text", "tekst; start; ile; nowy"], ["REPT", "text", "tekst; razy"], ["EXACT", "text", "tekst1; tekst2"],
    ["ISBLANK", "info", "wartość"], ["ISNUMBER", "info", "wartość"], ["ISTEXT", "info", "wartość"], ["ISERROR", "info", "wartość"], ["ISNA", "info", "wartość"],
  ];

  // ── tabele Excela: Tabela[Kolumna], Tabela[[#This Row],[od]], Tabela[@od], Tabela[] ──
  // tables: { nazwa(lower): { sheet, r1, c1, r2, c2, columns: [nazwy], headerRows, totalsRows } }
  function splitSpec(spec) {
    const items = [];
    let i = 0;
    const s = String(spec).trim();
    if (!s) return items;
    if (s[0] !== "[" && s[0] !== "@") return [s];
    if (s[0] === "@") {
      const rest = s.slice(1).trim();
      items.push("#This Row");
      if (rest) items.push(rest.replace(/^\[|\]$/g, ""));
      return items;
    }
    while (i < s.length) {
      if (s[i] === "[") {
        let depth = 0, j = i;
        for (; j < s.length; j++) {
          if (s[j] === "'" ) { j++; continue; }
          if (s[j] === "[") depth++;
          else if (s[j] === "]") { depth--; if (depth === 0) break; }
        }
        items.push(s.slice(i + 1, j));
        i = j + 1;
      } else if (s[i] === ":") { items.push(":"); i++; }
      else i++;
    }
    return items;
  }
  const unescapeCol = (name) => String(name).replace(/'(.)/g, "$1").trim().toLowerCase();
  function makeStructuredResolver(tables) {
    return (tableName, spec, ctx) => {
      const t = tables && tables[String(tableName).toLowerCase()];
      if (!t) return null;
      const items = splitSpec(spec);
      let rowSel = "#data";
      const cols = [];
      let colRange = false;
      for (const it of items) {
        if (it === ":") { colRange = true; continue; }
        const low = it.trim().toLowerCase();
        if (low === "#this row" || low === "#ten wiersz") rowSel = "#this";
        else if (low === "#headers" || low === "#nagłówki") rowSel = "#headers";
        else if (low === "#totals" || low === "#sumy") rowSel = "#totals";
        else if (low === "#all" || low === "#wszystko") rowSel = "#all";
        else if (low === "#data" || low === "#dane") rowSel = rowSel === "#headers" ? "#headers+data" : "#data";
        else cols.push(unescapeCol(it));
      }
      const idx = (name) => t.columns.findIndex((c) => c.toLowerCase() === name);
      let c1 = t.c1, c2 = t.c2;
      if (cols.length) {
        const a = idx(cols[0]);
        const b = colRange && cols[1] != null ? idx(cols[1]) : a;
        if (a < 0 || b < 0) return null;
        c1 = t.c1 + Math.min(a, b);
        c2 = t.c1 + Math.max(a, b);
      }
      const dataStart = t.r1 + (t.headerRows ?? 1);
      const dataEnd = t.r2 - (t.totalsRows || 0);
      let r1 = dataStart, r2 = dataEnd;
      if (rowSel === "#this") {
        // Excel: „ten wiersz” = ten sam NUMER wiersza, także z formuły w innym arkuszu.
        if (ctx.row < dataStart || ctx.row > dataEnd) return null;
        r1 = r2 = ctx.row;
      } else if (rowSel === "#headers") { r1 = t.r1; r2 = t.r1; }
      else if (rowSel === "#headers+data") { r1 = t.r1; }
      else if (rowSel === "#totals") { if (!t.totalsRows) return null; r1 = r2 = t.r2; }
      else if (rowSel === "#all") { r1 = t.r1; r2 = t.r2; }
      return { sheet: t.sheet, r1, c1, r2: Math.max(r1, r2), c2 };
    };
  }

  // Nazwy zdefiniowane (Formuły → Menedżer nazw): [{ Name, Ref, Sheet? }] z SheetJS.
  function makeNameResolver(names, sheetNames) {
    const cache = new Map();
    const list = Array.isArray(names) ? names : [];
    return (name, sheetScope, ctx) => {
      const low = String(name).toLowerCase();
      const scopeIdx = sheetNames.indexOf(sheetScope != null ? sheetScope : ctx.sheetName);
      const def = list.find((n) => n.Name && n.Name.toLowerCase() === low && n.Sheet === scopeIdx)
        || list.find((n) => n.Name && n.Name.toLowerCase() === low && (n.Sheet === undefined || n.Sheet === null));
      if (!def || !def.Ref || def.Ref === "#N/A") return null;
      let ast = cache.get(def);
      if (ast === undefined) {
        try { ast = parse(def.Ref, { lang: "en" }); } catch { ast = null; }
        cache.set(def, ast);
      }
      if (!ast) return null;
      if ((ctx.depth || 0) > 20) return ERR["#NAME?"];
      return evalNode(ast, { ...ctx, depth: (ctx.depth || 0) + 1 });
    };
  }

  const api = {
    makeStructuredResolver, makeNameResolver, makeMatrixCache, setLocale, roundExcel, excelPow,
    parse, serialize, localize, shift, syntaxLang, evaluate, refsOf, functionsOf, isSupported, resultLooksLikeDate,
    tokenize, colToIndex, indexToCol, quoteSheet, Range, isRange, isErr, ERR, ERR_PL, UNSUPPORTED,
    jsDateToSerial, dateSerial, generalText, PL_NAMES, EN_FROM_PL, CATALOG, makeCriterion, compare,
  };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  root.FormulaEngine = api;
})(typeof window !== "undefined" ? window : globalThis);
