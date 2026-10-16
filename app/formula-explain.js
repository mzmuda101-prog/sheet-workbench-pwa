// formula-explain.js — formuły zwykłymi słowami.
//
// „Wyjaśnij formułę” (menu komórki): drzewo formuły z silnika (formula-engine.js) → zdanie po
// polsku, z nazwami KOLUMN zamiast samych adresów („„Data od” (C5)”), aktualnymi wartościami
// odwołań i wynikiem. Bez AI — zawsze to samo i zawsze zgodne z tym, co liczy silnik.
//
// „✦ Opisz słowami” (okno fx): asystent AI (local-ai.js) proponuje formułę, a aplikacja ją
// SPRAWDZA zanim ją pokaże: składnia, znane funkcje, brak odwołania do samej komórki, wynik na
// aktywnym wierszu — i pokazuje jej wyjaśnienie z tego samego generatora, co wyżej. Wstawienie
// trafia tylko do paska formuły; zatwierdza człowiek (Enter / ✓), jak każdą formułę.

const SWB_EXPLAIN_OPS = {
  "+": "plus", "-": "minus", "*": "razy", "/": "podzielone przez", "^": "do potęgi", "&": "połączone z",
  "=": "równa się", "<>": "jest różne od", ">": "jest większe niż", ">=": "jest co najmniej", "<": "jest mniejsze niż", "<=": "jest co najwyżej",
};

function swbExplainFn(name) {
  const lang = typeof currentLang === "string" ? currentLang : "pl";
  return lang === "pl" && FormulaEngine.PL_NAMES[name] ? FormulaEngine.PL_NAMES[name] : name;
}

function swbExplainColLetter(c) {
  return FormulaEngine.indexToCol ? FormulaEngine.indexToCol(c) : XLSX.utils.encode_col(c);
}

// Nazwa kolumny tabeli dla bezwzględnego numeru kolumny (tylko bieżący arkusz)
function swbExplainHeader(sheet, c) {
  if (sheet && sheet !== currentSheetName) return "";
  const i = c - currentStartCol;
  return i >= 0 && i < currentHeaders.length ? String(currentHeaders[i] || "").trim() : "";
}

function swbExplainCriterion(node, describe) {
  if (node.t !== "str") return `warunek ${describe(node)}`;
  const m = String(node.v).match(/^(<>|>=|<=|=|>|<)?(.*)$/);
  const op = m[1] || "=";
  const val = m[2];
  const shown = /^-?\d+([.,]\d+)?$/.test(val) ? val : `„${val}”`;
  if (val === "" && op === "=") return "jest puste";
  if (val === "" && op === "<>") return "nie jest puste";
  if (/[*?]/.test(val) && op === "=") return `pasuje do wzorca ${shown}`;
  return `${SWB_EXPLAIN_OPS[op]} ${shown}`;
}

function swbDescribeNode(node, ctx) {
  const d = (n) => swbDescribeNode(n, ctx);
  switch (node.t) {
    case "num": return String(node.v).replace(".", ",");
    case "str": return node.v === "" ? "puste" : `„${node.v}”`;
    case "bool": return node.v ? "PRAWDA" : "FAŁSZ";
    case "err": return `błąd ${node.v}`;
    case "paren": return d(node.a);
    case "pct": return `${d(node.a)}%`;
    case "un": return node.op === "-" ? `minus ${d(node.a)}` : d(node.a);
    case "name": return `nazwa „${node.v}”`;
    case "ref": {
      const addr = `${swbExplainColLetter(node.c)}${node.r + 1}`;
      const h = swbExplainHeader(node.sheet, node.c);
      ctx.refs.push({ kind: "ref", sheet: node.sheet, r: node.r, c: node.c, addr, header: h });
      const where = node.sheet && node.sheet !== currentSheetName ? ` z arkusza „${node.sheet}”` : "";
      return h ? `„${h}” (${addr}${where})` : `${addr}${where}`;
    }
    case "range": {
      const whole = node.cols || node.r2 == null;
      const addr = whole ? `${swbExplainColLetter(node.c1)}:${swbExplainColLetter(node.c2)}` : `${swbExplainColLetter(node.c1)}${node.r1 + 1}:${swbExplainColLetter(node.c2)}${node.r2 + 1}`;
      const where = node.sheet && node.sheet !== currentSheetName ? ` w arkuszu „${node.sheet}”` : "";
      const h = node.c1 === node.c2 ? swbExplainHeader(node.sheet, node.c1) : "";
      ctx.refs.push({ kind: "range", addr: `${node.sheet ? `${node.sheet}!` : ""}${addr}`, header: h });
      if (h) return `kolumna „${h}” (${addr}${where})`;
      return whole && node.c1 === node.c2 ? `cała kolumna ${swbExplainColLetter(node.c1)}${where}` : `zakres ${addr}${where}`;
    }
    case "sref": {
      // tabela Excela: Tabela[[#This Row],[od]] / Tabela[@od] / Tabela[Kwota]
      const spec = String(node.spec || "");
      const thisRow = /#this row|^@|\[@/i.test(spec);
      const cols = [...spec.matchAll(/\[([^\[\]#][^\[\]]*)\]/g)].map((m) => m[1].replace(/^@/, ""));
      if (!cols.length && !/[\[\]#]/.test(spec)) cols.push(spec.replace(/^@/, ""));
      const col = cols.join(" : ");
      const addr = `${node.table}[${spec}]`;
      ctx.refs.push({ kind: thisRow ? "sref-row" : "sref", addr, header: col, text: addr });
      // nazwa tabeli jest na liście odwołań — w zdaniu wystarczy kolumna
      return col ? `„${col}”${thisRow ? "" : ` (cała kolumna tabeli „${node.table}”)`}` : `tabela „${node.table}”`;
    }
    case "array": return `tablica {${node.rows.map((r) => r.map(d).join("; ")).join(" | ")}}`;
    case "bin": {
      // X = "" → „X jest puste”, X <> "" → „X nie jest puste”
      const empty = (n) => n.t === "str" && n.v === "";
      if ((node.op === "=" || node.op === "<>") && (empty(node.a) || empty(node.b))) {
        const other = empty(node.a) ? node.b : node.a;
        return `${d(other)} ${node.op === "=" ? "jest puste" : "nie jest puste"}`;
      }
      return `${d(node.a)} ${SWB_EXPLAIN_OPS[node.op] || node.op} ${d(node.b)}`;
    }
    case "call": return swbDescribeCall(node, ctx, d);
    default: return "?";
  }
}

function swbDescribeCall(node, ctx, d) {
  const a = node.args;
  const fn = node.name;
  const list = (xs) => xs.map(d).join(", ");
  const pairs = (xs, start = 0) => {
    const out = [];
    for (let i = start; i + 1 < xs.length; i += 2) out.push(`${d(xs[i])} ${swbExplainCriterion(xs[i + 1], d)}`);
    return out.join(" oraz ");
  };
  const same = (x, y) => { try { return FormulaEngine.serialize(x) === FormulaEngine.serialize(y); } catch { return false; } };
  const isEmptyStr = (n) => n && n.t === "str" && n.v === "";
  switch (fn) {
    case "IF": {
      // JEŻELI(X>0; X&"m "; "") — częsty wzorzec „licznik z dopiskiem, gdy jest co pokazać”
      const cond = a[0];
      const then = a[1];
      if (cond?.t === "bin" && cond.op === ">" && cond.b?.t === "num" && cond.b.v === 0 && then?.t === "bin" && then.op === "&"
        && same(then.a, cond.a) && then.b?.t === "str" && isEmptyStr(a[2])) {
        return `${d(cond.a)} z dopiskiem „${then.b.v.trim()}” (gdy większe od 0)`;
      }
      return `jeżeli ${d(a[0])} — ${a[1] ? d(a[1]) : "PRAWDA"}, w przeciwnym razie ${a[2] ? d(a[2]) : "FAŁSZ"}`;
    }
    case "IFS": {
      const parts = [];
      for (let i = 0; i + 1 < a.length; i += 2) parts.push(`gdy ${d(a[i])} — ${d(a[i + 1])}`);
      return parts.join("; ");
    }
    case "IFERROR": return `${d(a[0])}; gdy wyjdzie błąd — ${d(a[1])}`;
    case "IFNA": return `${d(a[0])}; gdy brak wartości (#N/D) — ${d(a[1])}`;
    case "AND": return a.map(d).join(" i ");
    case "OR": return a.map(d).join(" lub ");
    case "NOT": return `nieprawda, że ${d(a[0])}`;
    case "SUM": return `suma ${list(a)}`;
    case "AVERAGE": return `średnia z ${list(a)}`;
    case "MIN": return `najmniejsza wartość z ${list(a)}`;
    case "MAX": return `największa wartość z ${list(a)}`;
    case "MEDIAN": return `mediana z ${list(a)}`;
    case "COUNT": return `liczba komórek z liczbami w ${list(a)}`;
    case "COUNTA": return `liczba niepustych komórek w ${list(a)}`;
    case "COUNTBLANK": return `liczba pustych komórek w ${list(a)}`;
    case "COUNTIF": return `ile komórek w ${d(a[0])} ${swbExplainCriterion(a[1], d)}`;
    case "COUNTIFS": return `ile wierszy spełnia jednocześnie: ${pairs(a)}`;
    case "SUMIF": return `suma ${d(a[2] || a[0])} tam, gdzie ${d(a[0])} ${swbExplainCriterion(a[1], d)}`;
    case "SUMIFS": return `suma ${d(a[0])} tam, gdzie ${pairs(a, 1)}`;
    case "AVERAGEIF": return `średnia ${d(a[2] || a[0])} tam, gdzie ${d(a[0])} ${swbExplainCriterion(a[1], d)}`;
    case "AVERAGEIFS": return `średnia ${d(a[0])} tam, gdzie ${pairs(a, 1)}`;
    case "MINIFS": return `najmniejsza wartość ${d(a[0])} tam, gdzie ${pairs(a, 1)}`;
    case "MAXIFS": return `największa wartość ${d(a[0])} tam, gdzie ${pairs(a, 1)}`;
    case "VLOOKUP": return `znajdź ${d(a[0])} w pierwszej kolumnie ${d(a[1])} i weź wartość z ${d(a[2])}. kolumny${a[3] && a[3].t === "bool" && !a[3].v || a[3]?.t === "num" && a[3].v === 0 ? " (dokładne dopasowanie)" : " (dopasowanie przybliżone)"}`;
    case "HLOOKUP": return `znajdź ${d(a[0])} w pierwszym wierszu ${d(a[1])} i weź wartość z ${d(a[2])}. wiersza`;
    case "XLOOKUP": return `znajdź ${d(a[0])} w ${d(a[1])} i zwróć odpowiadającą wartość z ${d(a[2])}${a[3] ? `; gdy nie znajdzie — ${d(a[3])}` : ""}`;
    case "MATCH": return `pozycja ${d(a[0])} w ${d(a[1])}`;
    case "INDEX": return `wartość z ${d(a[0])} w wierszu nr ${a[1] ? d(a[1]) : "?"}${a[2] ? ` i kolumnie nr ${d(a[2])}` : ""}`;
    case "TODAY": return "dzisiejsza data";
    case "NOW": return "bieżąca data i godzina";
    case "DATEDIF": {
      const unit = a[2]?.t === "str" ? String(a[2].v).toUpperCase() : "";
      const u = { D: "dni", M: "pełnych miesięcy", Y: "pełnych lat", MD: "dni (bez miesięcy)", YM: "miesięcy (bez lat)", YD: "dni (bez lat)" }[unit] || "jednostek";
      const until = a[1]?.t === "call" && a[1].name === "TODAY" ? "dziś" : d(a[1]);
      return `liczba ${u} od ${d(a[0])} do ${until}`;
    }
    case "DAYS": return `liczba dni od ${d(a[1])} do ${d(a[0])}`;
    case "NETWORKDAYS": return `liczba dni roboczych od ${d(a[0])} do ${d(a[1])}`;
    case "YEAR": return `rok z ${d(a[0])}`;
    case "MONTH": return `miesiąc z ${d(a[0])}`;
    case "DAY": return `dzień miesiąca z ${d(a[0])}`;
    case "DATE": return `data ${d(a[0])}-${d(a[1])}-${d(a[2])}`;
    case "EDATE": return `${d(a[0])} przesunięte o ${d(a[1])} mies.`;
    case "EOMONTH": return `ostatni dzień miesiąca dla ${d(a[0])} przesuniętego o ${d(a[1])} mies.`;
    case "CONCAT": case "CONCATENATE": return `tekst połączony z ${list(a)}`;
    case "TEXTJOIN": return `teksty z ${list(a.slice(2))} połączone separatorem ${d(a[0])}`;
    case "LEFT": return `pierwsze ${a[1] ? d(a[1]) : "1"} znaki z ${d(a[0])}`;
    case "RIGHT": return `ostatnie ${a[1] ? d(a[1]) : "1"} znaki z ${d(a[0])}`;
    case "MID": return `${d(a[2])} znaków z ${d(a[0])}, od znaku nr ${d(a[1])}`;
    case "LEN": return `długość tekstu ${d(a[0])}`;
    case "TRIM": return `${d(a[0])} bez zbędnych spacji`;
    case "UPPER": return `${d(a[0])} wielkimi literami`;
    case "LOWER": return `${d(a[0])} małymi literami`;
    case "TEXT": return `${d(a[0])} zapisane w formacie ${d(a[1])}`;
    case "VALUE": return `${d(a[0])} jako liczba`;
    case "ROUND": return `${d(a[0])} zaokrąglone do ${d(a[1])} miejsc po przecinku`;
    case "ROUNDUP": return `${d(a[0])} zaokrąglone w górę do ${d(a[1])} miejsc`;
    case "ROUNDDOWN": return `${d(a[0])} zaokrąglone w dół do ${d(a[1])} miejsc`;
    case "ABS": return `wartość bezwzględna ${d(a[0])}`;
    case "ISBLANK": return `${d(a[0])} jest puste`;
    case "ISNUMBER": return `${d(a[0])} jest liczbą`;
    case "ISTEXT": return `${d(a[0])} jest tekstem`;
    case "ISERROR": return `${d(a[0])} daje błąd`;
    default: return `${swbExplainFn(fn)}(${list(a)})`;
  }
}

function swbExplainValue(v) {
  if (v == null || v === "") return "(puste)";
  if (FormulaEngine.isErr && FormulaEngine.isErr(v)) return v.err;
  if (typeof v === "object" && v.unsupported) return "?";
  if (FormulaEngine.isRange && FormulaEngine.isRange(v)) return "(zakres)";
  if (typeof v === "boolean") return v ? "PRAWDA" : "FAŁSZ";
  if (typeof v === "number") return Number.isInteger(v) ? String(v) : String(Math.round(v * 1e6) / 1e6).replace(".", ",");
  return String(v);
}

// → { ok, sentence, refs: [{ addr, header, value }], result, unknown: [fn], selfRef }
function swbExplainFormula(text, { sheet = currentSheetName, r = 0, c = 0 } = {}) {
  const FE = FormulaEngine;
  const lang = typeof currentLang === "string" ? currentLang : "pl";
  let ast;
  try { ast = FE.parse(String(text).replace(/^\s*=?/, "="), { lang }); } catch (e) { return { ok: false, error: e.message }; }
  const ctx = { refs: [] };
  const cap = (x) => x.charAt(0).toLocaleUpperCase("pl") + x.slice(1);
  // Łańcuch JEŻELI(…; …; JEŻELI(…)) → kroki po kolei, zamiast jednego zdania na pół ekranu
  const steps = [];
  let node = ast;
  while (node.t === "call" && node.name === "IF" && node.args.length >= 2) {
    steps.push(`Jeżeli ${swbDescribeNode(node.args[0], ctx)} → ${swbDescribeNode(node.args[1], ctx)}`);
    node = node.args[2] || { t: "bool", v: false };
  }
  let sentence;
  if (steps.length >= 2) {
    steps.push(`W przeciwnym razie → ${swbDescribeNode(node, ctx)}`);
    sentence = "Formuła sprawdza warunki po kolei:";
  } else {
    steps.length = 0;
    ctx.refs.length = 0;
    sentence = cap(swbDescribeNode(ast, ctx)) + ".";
  }
  const unknown = [...FE.functionsOf(ast)].filter((n) => !FE.isSupported(n));
  const selfRef = FE.refsOf(ast).some((dref) => !dref.table && !dref.name && (dref.sheet == null || dref.sheet === sheet)
    && r >= Math.min(dref.r1, dref.r2) && r <= Math.max(dref.r1, dref.r2) && c >= Math.min(dref.c1, dref.c2) && c <= Math.max(dref.c1, dref.c2));
  let result = null;
  const evalCtx = typeof fcContext === "function" ? fcContext(sheet, r, c) : null;
  if (!unknown.length && !selfRef && evalCtx) {
    try { result = FE.evaluate(ast, evalCtx); } catch { result = null; }
  }
  const seen = new Set();
  const refs = ctx.refs.filter((x) => (seen.has(x.addr) ? false : seen.add(x.addr))).slice(0, 12).map((x) => {
    let value = null;
    const single = x.kind === "ref" || x.kind === "sref-row";
    if (single && evalCtx) {
      const src = x.kind === "ref" ? `=${x.sheet ? `${FE.quoteSheet(x.sheet)}!` : ""}${x.addr}` : `=${x.text}`;
      try { value = FE.evaluate(FE.parse(src, { lang: "en" }), evalCtx); } catch { value = null; }
    }
    // liczba w zakresie dat (1970–2100) w kolumnie o „datowej” nazwie → data, nie numer seryjny
    const dateLike = typeof value === "number" && value > 25569 && value < 73051 && /dat|^od$|^do$|start|koniec|termin|date|from|to$/i.test(x.header || "");
    const shown = dateLike ? new Date(Math.round((value - 25569) * 86400000)).toISOString().slice(0, 10) : swbExplainValue(value);
    return { addr: x.addr, header: x.kind.startsWith("sref") ? "" : x.header, value: single ? shown : null };
  });
  return { ok: true, ast, sentence, steps, refs, result: result == null ? null : swbExplainValue(result), unknown, selfRef };
}

// ── okno ─────────────────────────────────────────────────────────────────────
let swbExplainDialog = null;
function swbExplainEnsureDialog() {
  if (swbExplainDialog) return swbExplainDialog;
  swbExplainDialog = document.createElement("dialog");
  swbExplainDialog.className = "lai-dialog fx-explain-dialog";
  swbExplainDialog.setAttribute("aria-labelledby", "fxExplainTitle");
  swbExplainDialog.addEventListener("click", (e) => { if (e.target === swbExplainDialog) swbExplainDialog.close(); });
  document.body.appendChild(swbExplainDialog);
  return swbExplainDialog;
}

function swbExplainRender(box, formula, info) {
  const el = (tag, cls, text) => Object.assign(document.createElement(tag), { className: cls || "", textContent: text ?? "" });
  box.append(el("code", "fx-explain-formula", formula));
  if (!info.ok) { box.append(el("p", "lai-warn", t("fxExplainSyntax", { msg: info.error }))); return; }
  box.append(el("p", "fx-explain-sentence", info.sentence));
  if (info.steps?.length) {
    const ol = el("ol", "fx-explain-steps");
    info.steps.forEach((x) => ol.append(el("li", "", x)));
    box.append(ol);
  }
  if (info.refs.length) {
    const ul = el("ul", "fx-explain-refs");
    info.refs.forEach((x) => ul.append(el("li", "", `${x.addr}${x.header ? ` — „${x.header}”` : ""}${x.value != null ? ` = ${x.value}` : ""}`)));
    box.append(el("div", "fx-explain-label", t("fxExplainRefs")), ul);
  }
  if (info.unknown.length) box.append(el("p", "lai-note", t("fxExplainUnknown", { fns: info.unknown.map(swbExplainFn).join(", ") })));
  if (info.selfRef) box.append(el("p", "lai-warn", t("fxExplainSelfRef")));
  if (info.result != null) box.append(el("p", "fx-explain-result", t("fxExplainResult", { value: info.result })));
}

function swbOpenFormulaExplain({ formula, sheet = currentSheetName, r, c }) {
  const dlg = swbExplainEnsureDialog();
  dlg.replaceChildren();
  const box = Object.assign(document.createElement("div"), { className: "lai-box" });
  const title = Object.assign(document.createElement("h2"), { className: "lai-title", id: "fxExplainTitle", textContent: t("fxExplainTitle") });
  box.append(title);
  swbExplainRender(box, formula, swbExplainFormula(formula, { sheet, r, c }));
  const close = Object.assign(document.createElement("button"), { type: "button", className: "lai-btn", textContent: t("fxExplainClose") });
  close.addEventListener("click", () => dlg.close());
  const row = Object.assign(document.createElement("div"), { className: "lai-actions" });
  row.append(close);
  box.append(row);
  dlg.append(box);
  if (!dlg.open) dlg.showModal();
}

// ── ✦ Opisz słowami → formuła (AI + sprawdzenie) ─────────────────────────────
function swbFormulaAiSystem(target) {
  const cols = currentHeaders.slice(0, 40).map((h, i) => `${swbExplainColLetter(currentStartCol + i)} = „${h}”`).join("; ");
  return "Piszesz JEDNĄ formułę Excela do komórki " + target.ref + " (wiersz " + (target.r + 1) + "). " +
    "Kolumny arkusza: " + cols + ". Odwołuj się do komórek z tego samego wiersza (np. C" + (target.r + 1) + "). " +
    "Używaj angielskich nazw funkcji i przecinków między argumentami. Odpowiedz WYŁĄCZNIE formułą zaczynającą się od =, bez komentarza.";
}

function swbFormulaFromAnswer(answer) {
  const s = String(answer || "").replace(/```[a-z]*|```/gi, "").trim();
  const line = s.split(/\n/).map((x) => x.trim().replace(/^`|`$/g, "")).find((x) => x.startsWith("=")) || "";
  return line.replace(/[.;]\s*$/, "");
}

// target: { r, c, ref } aktywnej komórki → { formula, info } albo { error }
async function swbWriteFormula(description, target) {
  const answer = await LocalAI.generate({ system: swbFormulaAiSystem(target), user: description, maxTokens: 120 });
  const formula = swbFormulaFromAnswer(answer);
  if (!formula) return { error: t("fxWriteNoFormula") };
  const info = swbExplainFormula(formula, { sheet: currentSheetName, r: target.r, c: target.c });
  if (!info.ok) return { error: t("fxExplainSyntax", { msg: info.error }), formula };
  if (info.unknown.length) return { error: t("fxExplainUnknown", { fns: info.unknown.join(", ") }), formula };
  if (info.selfRef) return { error: t("fxExplainSelfRef"), formula };
  const lang = typeof currentLang === "string" ? currentLang : "pl";
  return { formula: "=" + FormulaEngine.serialize(info.ast, { lang }), info };
}

function swbFormulaTarget() {
  const a = typeof cfmtActiveCell === "function" ? cfmtActiveCell() : null;
  if (!a) return null;
  return { r: a.row.rowIndex0, c: currentStartCol + a.col, ref: a.ref };
}

// Sekcja w oknie fx (formula-ui.js wywołuje przy budowie okna)
function swbFormulaWriterSection(onInsert) {
  const wrap = Object.assign(document.createElement("div"), { className: "fx-write" });
  const row = Object.assign(document.createElement("div"), { className: "fx-write-row" });
  const input = Object.assign(document.createElement("input"), { type: "text", className: "fx-write-input", placeholder: t("fxWritePlaceholder") });
  input.setAttribute("aria-label", t("fxWritePlaceholder"));
  const go = Object.assign(document.createElement("button"), { type: "button", className: "btn btn-sm", textContent: t("fxWriteGo") });
  row.append(Object.assign(document.createElement("span"), { className: "nlq-icon", textContent: "✦", ariaHidden: "true" }), input, go);
  const out = Object.assign(document.createElement("div"), { className: "fx-write-out" });
  out.hidden = true;
  const run = async () => {
    const desc = input.value.trim();
    if (!desc) { input.focus(); return; }
    const target = swbFormulaTarget();
    if (!target) { toast(t("fxWriteNoCell"), "info"); return; }
    if (!(await LocalAI.ensure({ reason: t("fxWriteReason") }))) return;
    go.disabled = true;
    out.hidden = false;
    out.replaceChildren(Object.assign(document.createElement("p"), { className: "lai-note", textContent: "…" }));
    try {
      const res = await swbWriteFormula(desc, target);
      out.replaceChildren(Object.assign(document.createElement("span"), { className: "ai-badge", textContent: t("aiBadge") }));
      if (res.error) {
        if (res.formula) out.append(Object.assign(document.createElement("code"), { className: "fx-explain-formula", textContent: res.formula }));
        out.append(Object.assign(document.createElement("p"), { className: "lai-warn", textContent: res.error }));
        return;
      }
      swbExplainRender(out, res.formula, res.info);
      const ins = Object.assign(document.createElement("button"), { type: "button", className: "btn btn-sm", textContent: t("fxWriteInsert") });
      ins.addEventListener("click", () => onInsert(res.formula));
      out.append(ins);
    } catch (err) {
      out.replaceChildren(Object.assign(document.createElement("p"), { className: "lai-warn", textContent: t("aiFailed", { msg: err.message || err }) }));
    } finally {
      go.disabled = false;
    }
  };
  go.addEventListener("click", run);
  input.addEventListener("keydown", (e) => { if (e.key === "Enter" && !e.isComposing) { e.preventDefault(); e.stopPropagation(); run(); } });
  wrap.append(row, out);
  return wrap;
}
