// anonymize-tools.js — anonimizacja, etapy 3–4 (2026-11-06). Bez DOM — ładowane też w Node (testy).
//
// Etap 3: polecenie tekstem („PESEL zamaskuj, nazwiska na pseudonimy, daty do miesiąca, notatki
//         usuń”) → ustawienia okna; profile (wbudowane + własne, import/eksport pliku JSON).
// Etap 4: raport ryzyka (k-anonimowość: ile wierszy wciąż ma UNIKALNĄ kombinację cech typu
//         kod pocztowy + data urodzenia + płeć) i klucz odwracania — lista zmian komórek
//         zaszyfrowana hasłem (PBKDF2-SHA256 → AES-GCM), nic nie wychodzi z urządzenia.
//
// Wszystko tu tylko PROPONUJE ustawienia albo liczy — zmiany w pliku robi anonymize-ui.js.

const SwbAnonTools = (() => {
  const norm = (s) => String(s ?? "").normalize("NFC").trim().replace(/\s+/g, " ").toLocaleLowerCase("pl");
  const fold = (s) => norm(s).normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/ł/g, "l");
  const W = "(?<![a-z0-9])"; // początek słowa (po fold() są tylko litery ASCII)
  const E = "(?![a-z0-9])";
  const re = (src) => new RegExp(src, "g");

  const TYPES = ["custom", "person", "pesel", "nip", "regon", "iban", "card", "idcard", "passport", "kw", "email", "phone", "address", "postcode", "place", "dob", "plate", "vin", "ip", "other"];
  const METHODS = ["keep", "pseudo", "fake", "mask", "generalize", "redact", "remove", "hash", "shift"];

  // ── Etap 3: polecenie tekstem ───────────────────────────────────────────
  // Kolejność ma znaczenie: dopasowany fragment jest wycinany, więc „adres e-mail” nie trafi
  // potem w „adres”, a „kod pocztowy” w „kod”.
  const TYPE_WORDS = [
    ["email", re(`${W}(adres\\w*\\s+)?(e-?mail\\w*|maile?\\w*|poczt\\w*\\s+elektroniczn\\w*)${E}`)],
    ["ip", re(`${W}(adres\\w*\\s+)?ip${E}`)],
    ["postcode", re(`${W}(kod\\w*\\s+pocztow\\w*|kod\\w*|postcode\\w*|zip\\s*code\\w*|postal\\s+code\\w*)${E}`)],
    ["dob", re(`${W}(dat\\w*(\\s+urodzeni\\w*)?|urodzin\\w*|dob|dates?(\\s+of\\s+birth)?|birth\\s*dates?|birthdays?)${E}`)],
    ["pesel", re(`${W}pesel\\w*${E}`)],
    ["nip", re(`${W}(nip\\w*|tax\\s+ids?)${E}`)],
    ["regon", re(`${W}regon\\w*${E}`)],
    ["iban", re(`${W}((numer\\w*|nr)\\s+)?(kont(o|a|ach|om|ami)?|rachunk\\w*|rachunek|iban\\w*|nrb|accounts?|bank\\s+accounts?)${E}`)],
    ["card", re(`${W}(kart\\w*(\\s+platnicz\\w*)?|cards?|payment\\s+cards?)${E}`)],
    ["idcard", re(`${W}(dowod\\w*(\\s+osobist\\w*)?|id\\s+cards?)${E}`)],
    ["passport", re(`${W}(paszport\\w*|passports?)${E}`)],
    ["kw", re(`${W}(ksieg\\w*(\\s+wieczyst\\w*)?|kw|land\\s+registers?)${E}`)],
    ["phone", re(`${W}((numer\\w*|nr)\\s+)?(telefon\\w*|tel|komork\\w*|phones?|phone\\s+numbers?)${E}`)],
    ["address", re(`${W}(adres\\w*|ulic\\w*|address\\w*|streets?)${E}`)],
    ["place", re(`${W}(miejscowos\\w*|miast\\w*|miejsc\\w*\\s+zamieszkani\\w*|towns?|cit(y|ies))${E}`)],
    ["plate", re(`${W}(rejestracj\\w*|tablic\\w*|plates?|license\\s+plates?)${E}`)],
    ["vin", re(`${W}vin\\w*${E}`)],
    ["person", re(`${W}(osob\\w*|nazwisk\\w*|imion\\w*|imie|imienia|ludzi|pracownik\\w*|klient\\w*|wlasciciel\\w*|personalia|people|persons?|names?|surnames?|employees?|customers?|owners?)${E}`)],
    ["custom", re(`${W}(wlasn\\w*\\s+(slow\\w*|reguł\\w*|regul\\w*|fraz\\w*)|custom\\s+(words?|rules?|terms?))${E}`)],
  ];
  const ALL_WORDS = re(`${W}(wszystk\\w*|cal(e|a|y)\\s+dane|everything|all)${E}`);
  const REST_WORDS = re(`${W}(reszt\\w*|pozostal\\w*|inne\\s+dane|rest|others?|remaining)${E}`);

  // [metoda, wzorzec]; „nie ruszaj” przed resztą (zawiera czasownik).
  const METHOD_WORDS = [
    ["keep", re(`${W}(zostaw\\w*|nie\\s+ruszaj\\w*|nie\\s+zmieniaj\\w*|bez\\s+zmian|pomin\\w*|zachowaj\\w*|keep|leave|skip|don'?t\\s+touch|unchanged)${E}`)],
    ["pseudo", re(`${W}(pseudonim\\w*|pseudoni\\w*|pseudonyms?|pseudonymi[sz]\\w*)${E}`)],
    ["fake", re(`${W}(falszyw\\w*|fikcyjn\\w*|zmyslon\\w*|realistyczn\\w*|losow\\w*|wymysl\\w*|fake|fictional|realistic|random)${E}`)],
    ["mask", re(`${W}(zamaskuj\\w*|maskuj\\w*|mask\\w*|zakryj\\w*|gwiazdk\\w*|zagwiazdkuj\\w*)${E}`)],
    ["generalize", re(`${W}(uogoln\\w*|zaokragl\\w*|skroc\\w*|generali[sz]\\w*|round\\w*|do\\s+(roku|lat|miesiac\\w*)|tylko\\s+(rok|miesiac)|to\\s+(the\\s+)?(year|month)|(year|month)\\s+only)${E}`)],
    ["redact", re(`${W}(znacznik\\w*|oznacz\\w*|zaczern\\w*|ukryj\\w*|redact\\w*|tag\\w*|hide)${E}`)],
    ["remove", re(`${W}(usun\\w*|wyczysc\\w*|skasuj\\w*|wytn\\w*|wymaz\\w*|wyrzuc\\w*|remove\\w*|delete\\w*|clear\\w*|erase\\w*)${E}`)],
    ["hash", re(`${W}(skrot\\w*|hash\\w*|haszuj\\w*|zahaszuj\\w*)${E}`)],
    ["shift", re(`${W}(przesun\\w*|shift\\w*)${E}`)],
  ];
  const MONTH_RE = re(`${W}(miesiac\\w*|month\\w*)${E}`);
  const YEAR_RE = re(`${W}(rok|roku|lat|year\\w*)${E}`);

  // Ślady poza komórkami (etap 2) i wynik — rozpoznawane przed rodzajami danych.
  const TRACE_WORDS = [
    ["notes", re(`${W}(notat\\w*|komentarz\\w*|notes?|comments?)${E}`)],
    ["metadata", re(`${W}(metadan\\w*|wlasciwosc\\w*\\s+pliku|autor\\w*(\\s+pliku)?|metadata|properties|authors?)${E}`)],
    ["headersFooters", re(`${W}(naglowk\\w*\\s+i\\s+stopk\\w*|naglowk\\w*\\s+wydruku|stopk\\w*|headers?\\s+and\\s+footers?|footers?)${E}`)],
  ];
  const SCOPE_WORDS = [
    ["file", re(`${W}(cal(y|ym)\\s+plik\\w*|wszystki\\w*\\s+arkusz\\w*|whole\\s+file|all\\s+sheets)${E}`)],
    ["selection", re(`${W}(zaznaczeni\\w*|selection)${E}`)],
    ["sheet", re(`${W}(ten\\s+arkusz|tylko\\s+arkusz|biezacy\\s+arkusz|this\\s+sheet|current\\s+sheet)${E}`)],
  ];
  const OUTPUT_WORDS = [
    ["copy", re(`${W}(kopi\\w*|nowy\\s+plik|osobny\\s+plik|copy|new\\s+file)${E}`)],
    ["inplace", re(`${W}(w\\s+(otwartym\\s+)?pliku|w\\s+miejscu|nadpisz\\w*|in\\s+place|this\\s+file|open\\s+file)${E}`)],
  ];
  const COLUMN_RE = re(`${W}(kolumn\\w*|column\\w*)${E}`);

  // Zdejmuje dopasowania z tekstu (zastępuje spacjami — pozycje reszty zostają).
  function take(state, pattern) {
    const out = [];
    pattern.lastIndex = 0;
    state.s = state.s.replace(pattern, (m, ...rest) => { out.push({ text: m, index: rest[rest.length - 2] }); return " ".repeat(m.length); });
    return out;
  }
  const leftover = (s) => s.replace(/[^a-z0-9]+/g, " ").trim();
  const FILLER = new Set("i oraz a na do z ze w we jako dla na oraz tez rowniez prosze please and to as with for the a an by of in into".split(" "));
  const meaningful = (s) => leftover(s).split(" ").filter((w) => w && !FILLER.has(w) && w.length > 1);

  // headers: [{ key, header }] — kolumny z ostatniego skanu (dla „kolumnę Uwagi usuń”).
  // Zwraca { methods, traces, output, scope, dateGeneralize, columns:[{key, header, mode, method?}],
  //          understood:[{kind, ...}], unknown:[fragment] }.
  function parseCommand(text, opts = {}) {
    const headers = (opts.headers || []).map((h) => ({ ...h, f: fold(h.header) })).filter((h) => h.f.length >= 2);
    const res = { methods: {}, traces: {}, output: null, scope: null, dateGeneralize: null, columns: [], understood: [], unknown: [] };
    const original = String(text || "");
    const clauses = original.split(/[,;\n]+|\s+(?:oraz|a\s+także|a\s+takze|then|and\s+then)\s+|\.(?=\s|$)/i).map((x) => x.trim()).filter(Boolean);
    let pendingTypes = []; // rodzaje bez metody czekają na następną klauzulę z metodą
    let pendingRaw = [];
    const explicitTypes = new Set();
    const setMethod = (type, method, raw) => {
      if (!TYPES.includes(type) || !METHODS.includes(method)) return;
      // przesunięcie dat ma sens tylko dla dat i własnych reguł
      if (method === "shift" && !["dob", "custom", "other"].includes(type)) return;
      res.methods[type] = method;
      explicitTypes.add(type);
      res.understood.push({ kind: "method", type, method, raw });
    };
    for (const raw of clauses) {
      const state = { s: fold(raw) };
      // 1. kolumna („kolumnę Uwagi usuń”, „kolumna "Telefon" zostaw”)
      if (headers.length && (COLUMN_RE.test(state.s) || /["„”'«»]/.test(raw))) {
        COLUMN_RE.lastIndex = 0;
        const quoted = (raw.match(/["„”'«»]([^"„”'«»]+)["„”'«»]/) || [])[1];
        const target = quoted ? fold(quoted) : null;
        const hit = headers
          .filter((h) => (target ? h.f === target : new RegExp(`${W}${h.f.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}${E}`).test(state.s)))
          .sort((a, b) => b.f.length - a.f.length)[0];
        if (hit) {
          state.s = state.s.replace(hit.f, " ");
          if (quoted) state.s = state.s.replace(fold(quoted), " ");
          take(state, COLUMN_RE);
          const m = METHOD_WORDS.find(([, p]) => { p.lastIndex = 0; return p.test(state.s); });
          let type = null;
          for (const [t, p] of TYPE_WORDS) { p.lastIndex = 0; if (p.test(state.s)) { type = t; break; } }
          if (m && m[0] === "keep") res.columns.push({ key: hit.key, header: hit.header, mode: "skip" });
          else if (type) {
            res.columns.push({ key: hit.key, header: hit.header, mode: type });
            if (m) setMethod(type, m[0], raw);
          } else if (m) res.columns.push({ key: hit.key, header: hit.header, mode: "custom", method: m[0] });
          else { res.unknown.push(raw); continue; }
          res.understood.push({ kind: "column", header: hit.header, mode: res.columns[res.columns.length - 1].mode, method: res.columns[res.columns.length - 1].method || (m && m[0]) || null, raw });
          continue;
        }
      }
      // 2. zakres i wynik
      let used = false;
      for (const [scope, p] of SCOPE_WORDS) if (take(state, p).length) { res.scope = scope; res.understood.push({ kind: "scope", scope, raw }); used = true; break; }
      for (const [output, p] of OUTPUT_WORDS) if (take(state, p).length) { res.output = output; res.understood.push({ kind: "output", output, raw }); used = true; break; }
      // 3. ślady (notatki, metadane, nagłówki/stopki)
      const traceHits = TRACE_WORDS.filter(([, p]) => take(state, p).length).map(([k]) => k);
      const methodHits = [];
      for (const [m, p] of METHOD_WORDS) { const got = take(state, p); if (got.length) methodHits.push({ m, index: got[0].index }); }
      methodHits.sort((a, b) => a.index - b.index);
      const method = methodHits.length ? methodHits[0].m : null;
      if (traceHits.length) {
        traceHits.forEach((k) => {
          let v;
          if (k === "metadata") v = method === "keep" ? false : true;
          else if (k === "notes") v = method === "keep" ? "keep" : method === "remove" ? "remove" : "anonymize";
          else v = method === "keep" ? "keep" : method === "remove" || method === "redact" ? "clear" : "anonymize";
          res.traces[k] = v;
          res.understood.push({ kind: "trace", trace: k, value: v, raw });
        });
        continue;
      }
      // 4. rodzaje danych
      const types = [];
      for (const [t, p] of TYPE_WORDS) if (take(state, p).length && !types.includes(t)) types.push(t);
      const all = take(state, ALL_WORDS).length > 0;
      const rest = take(state, REST_WORDS).length > 0;
      // „daty do miesiąca” / „do roku” — dokładność uogólniania dat
      if (method === "generalize" && (types.includes("dob") || all || rest)) {
        const f = fold(raw);
        MONTH_RE.lastIndex = 0; YEAR_RE.lastIndex = 0;
        if (MONTH_RE.test(f)) res.dateGeneralize = "month";
        else if (YEAR_RE.test(f)) res.dateGeneralize = "year";
      }
      if (!method) {
        if (types.length) { pendingTypes.push(...types); pendingRaw.push(raw); continue; }
        if (!used && meaningful(state.s).length) res.unknown.push(raw);
        continue;
      }
      const targets = [...pendingTypes, ...types];
      pendingTypes = []; pendingRaw = [];
      if (all) TYPES.forEach((t) => { if (t !== "other") targets.push(t); });
      if (rest) TYPES.forEach((t) => { if (t !== "other" && !explicitTypes.has(t) && !targets.includes(t)) targets.push(t); });
      if (!targets.length) { res.unknown.push(raw); continue; }
      [...new Set(targets)].forEach((t) => setMethod(t, method, raw));
    }
    if (pendingRaw.length) res.unknown.push(...pendingRaw);
    return res;
  }

  // ── Etap 3: profile ─────────────────────────────────────────────────────
  const TRACE_KEYS = { metadata: "bool", notes: ["anonymize", "remove", "keep"], headersFooters: ["anonymize", "clear", "keep"], formulaCache: "bool", charts: "bool", pivots: "bool", links: "bool", sharedStrings: "bool" };
  const ALL_TRACES_ON = { metadata: true, formulaCache: true, charts: true, pivots: true, links: true, sharedStrings: true };
  // Wbudowane: nazwy i opisy w UI (pr_<id>, prd_<id>); tu same ustawienia.
  const PRESETS = [
    {
      id: "rodo-share", builtin: true,
      methods: { person: "pseudo", pesel: "mask", nip: "mask", regon: "mask", iban: "mask", card: "mask", idcard: "mask", passport: "mask", kw: "mask", email: "mask", phone: "mask", address: "generalize", postcode: "generalize", place: "keep", dob: "generalize", plate: "mask", vin: "mask", ip: "mask", custom: "redact", other: "redact" },
      traces: { ...ALL_TRACES_ON, notes: "remove", headersFooters: "clear" }, output: "copy", dateGeneralize: "year",
    },
    {
      id: "test-data", builtin: true,
      methods: { person: "fake", pesel: "fake", nip: "fake", regon: "fake", iban: "fake", card: "fake", idcard: "fake", passport: "fake", kw: "fake", email: "fake", phone: "fake", address: "fake", postcode: "fake", place: "fake", dob: "shift", plate: "fake", vin: "fake", ip: "fake", custom: "fake", other: "fake" },
      traces: { ...ALL_TRACES_ON, notes: "anonymize", headersFooters: "anonymize" }, output: "copy", dateGeneralize: null,
    },
    {
      id: "max", builtin: true,
      methods: { person: "redact", pesel: "remove", nip: "redact", regon: "redact", iban: "remove", card: "remove", idcard: "remove", passport: "remove", kw: "redact", email: "redact", phone: "remove", address: "redact", postcode: "generalize", place: "generalize", dob: "generalize", plate: "redact", vin: "redact", ip: "redact", custom: "redact", other: "redact" },
      traces: { ...ALL_TRACES_ON, notes: "remove", headersFooters: "clear" }, output: "copy", dateGeneralize: "year",
    },
  ];

  // Oczyszcza profil z pliku/pamięci: tylko znane klucze i wartości. Zwraca profil albo null.
  // Hasła (klucza spójności) profil NIGDY nie zawiera.
  function sanitizeProfile(p) {
    if (!p || typeof p !== "object") return null;
    const name = String(p.name ?? "").trim().slice(0, 80);
    if (!name && !p.builtin) return null;
    const out = { name, methods: {}, traces: {}, rules: { terms: "", exceptions: "", patterns: "" }, columns: {} };
    if (p.methods && typeof p.methods === "object") for (const [t, m] of Object.entries(p.methods)) if (TYPES.includes(t) && METHODS.includes(m)) out.methods[t] = m;
    if (p.traces && typeof p.traces === "object") for (const [k, v] of Object.entries(p.traces)) {
      const spec = TRACE_KEYS[k];
      if (spec === "bool" && typeof v === "boolean") out.traces[k] = v;
      else if (Array.isArray(spec) && spec.includes(v)) out.traces[k] = v;
    }
    if (p.rules && typeof p.rules === "object") for (const k of ["terms", "exceptions", "patterns"]) if (typeof p.rules[k] === "string") out.rules[k] = p.rules[k].slice(0, 20000);
    if (p.columns && typeof p.columns === "object") for (const [h, m] of Object.entries(p.columns)) {
      const key = fold(h).slice(0, 200);
      if (key && (m === "skip" || m === "auto" || TYPES.includes(m))) out.columns[key] = m;
    }
    if (p.output === "copy" || p.output === "inplace") out.output = p.output;
    if (["file", "sheet", "columns", "selection"].includes(p.scope)) out.scope = p.scope;
    if (p.dateGeneralize === "month" || p.dateGeneralize === "year") out.dateGeneralize = p.dateGeneralize;
    if (typeof p.everything === "boolean") out.everything = p.everything;
    if (typeof p.command === "string") out.command = p.command.slice(0, 2000);
    return out;
  }
  const PROFILE_FORMAT = "swb-anon-profiles";
  function exportProfiles(list) {
    return JSON.stringify({ format: PROFILE_FORMAT, version: 1, profiles: list.map(sanitizeProfile).filter(Boolean) }, null, 2);
  }
  // Zwraca { profiles, skipped } albo rzuca Error("badFile").
  function importProfiles(text) {
    let data;
    try { data = JSON.parse(String(text)); } catch { throw new Error("badFile"); }
    const raw = Array.isArray(data) ? data : data && data.format === PROFILE_FORMAT && Array.isArray(data.profiles) ? data.profiles : data && data.name ? [data] : null;
    if (!raw) throw new Error("badFile");
    const profiles = [];
    let skipped = 0;
    raw.slice(0, 200).forEach((p) => { const s = sanitizeProfile({ ...p, builtin: false }); if (s && s.name) profiles.push(s); else skipped++; });
    return { profiles, skipped };
  }

  // ── Etap 4: raport ryzyka (k-anonimowość) ───────────────────────────────
  // Cechy pośrednie (quasi-identyfikatory): same nie wskazują osoby, ale razem często tak.
  const QI_HEADER_RE = /(^|[^a-z])(plec|gender|sex|wiek|age|rok\s+urodz|urodz|birth|stanowisk|zawod|position|job|narodowo|nationality|obywatel|stan\s+cywiln|marital|wyksztalc|education|dzial|department|kod\s+poczt|postcode|zip|miejscowo|miasto|city|town|gmina|powiat|wojewodztw)/;
  const QI_TYPES = new Set(["dob", "postcode", "place", "address"]);
  function isQuasiColumn(header, type) {
    if (QI_TYPES.has(type)) return true;
    return QI_HEADER_RE.test(fold(header));
  }
  const keyOf = (v) => (v == null ? "" : v instanceof Date ? v.toISOString().slice(0, 10) : fold(v));
  // records: tablica wierszy (tablice wartości, w kolejności qiNames). Puste wiersze (same puste
  // wartości) pomijane. Zwraca { total, unique, small, minK, columns, drop:[{name, unique}], examples }.
  function riskReport(records, qiNames, opts = {}) {
    const small = opts.smallK ?? 3;
    const rows = records.filter((rec) => rec.some((v) => keyOf(v) !== ""));
    const count = (cols) => {
      const m = new Map();
      rows.forEach((rec) => { const k = cols.map((i) => keyOf(rec[i])).join("\u0001"); m.set(k, (m.get(k) || 0) + 1); });
      return m;
    };
    const idx = qiNames.map((_, i) => i);
    const groups = count(idx);
    let unique = 0, smallN = 0, minK = Infinity;
    const examples = [];
    rows.forEach((rec) => {
      const k = groups.get(idx.map((i) => keyOf(rec[i])).join("\u0001"));
      if (k < minK) minK = k;
      if (k === 1) { unique++; if (examples.length < 5) examples.push(rec.map((v) => (v instanceof Date ? v.toISOString().slice(0, 10) : v == null ? "" : String(v)))); }
      if (k <= small) smallN++;
    });
    // która cecha najbardziej odpowiada za unikalność: liczba unikalnych po jej pominięciu
    const drop = qiNames.length > 1 ? qiNames.map((name, j) => {
      const cols = idx.filter((i) => i !== j);
      const g = count(cols);
      let u = 0;
      rows.forEach((rec) => { if (g.get(cols.map((i) => keyOf(rec[i])).join("\u0001")) === 1) u++; });
      return { name, unique: u };
    }).sort((a, b) => a.unique - b.unique) : [];
    return { total: rows.length, unique, small: smallN, smallK: small, minK: rows.length ? minK : 0, columns: qiNames, drop, examples };
  }

  // ── Etap 4: klucz odwracania ────────────────────────────────────────────
  const KEY_FORMAT = "swb-anon-key";
  const KDF_ITER = 310000;
  const subtle = () => (globalThis.crypto && globalThis.crypto.subtle) || null;
  const b64 = (u8) => { let s = ""; for (let i = 0; i < u8.length; i++) s += String.fromCharCode(u8[i]); return btoa(s); };
  const unb64 = (s) => Uint8Array.from(atob(String(s)), (c) => c.charCodeAt(0));
  const encVal = (v) => (v == null ? { t: "0" } : v instanceof Date ? { t: "d", v: v.getTime() } : typeof v === "number" ? { t: "n", v } : typeof v === "boolean" ? { t: "b", v } : { t: "s", v: String(v) });
  const decVal = (o) => (!o || o.t === "0" ? null : o.t === "d" ? new Date(o.v) : o.t === "n" ? Number(o.v) : o.t === "b" ? !!o.v : String(o.v));
  async function deriveKey(password, salt, iter) {
    const base = await subtle().importKey("raw", new TextEncoder().encode(String(password)), "PBKDF2", false, ["deriveKey"]);
    return subtle().deriveKey({ name: "PBKDF2", salt, iterations: iter, hash: "SHA-256" }, base, { name: "AES-GCM", length: 256 }, false, ["encrypt", "decrypt"]);
  }
  // changes: [{ sheet, r, c, before, after }] (after = wartość po anonimizacji albo null = pusta).
  // Zwraca tekst JSON pliku klucza. Bez hasła — błąd (klucz bez hasła byłby wyciekiem).
  async function encryptReverseKey(changes, password, meta = {}) {
    if (!subtle()) throw new Error("noCrypto");
    if (!password || String(password).length < 8) throw new Error("weakPassword");
    const payload = {
      file: meta.file || "", created: meta.created || new Date().toISOString(),
      changes: changes.map((ch) => ({ s: ch.sheet, r: ch.r, c: ch.c, b: encVal(ch.before), a: encVal(ch.after) })),
    };
    const salt = globalThis.crypto.getRandomValues(new Uint8Array(16));
    const iv = globalThis.crypto.getRandomValues(new Uint8Array(12));
    const key = await deriveKey(password, salt, KDF_ITER);
    const ct = new Uint8Array(await subtle().encrypt({ name: "AES-GCM", iv }, key, new TextEncoder().encode(JSON.stringify(payload))));
    return JSON.stringify({ format: KEY_FORMAT, version: 1, kdf: "PBKDF2-SHA256", iter: KDF_ITER, salt: b64(salt), iv: b64(iv), data: b64(ct), cells: changes.length });
  }
  // Zwraca { file, created, changes:[{sheet,r,c,before,after}] }; błędy: badFile, badPassword, noCrypto.
  async function decryptReverseKey(text, password) {
    if (!subtle()) throw new Error("noCrypto");
    let d;
    try { d = JSON.parse(String(text)); } catch { throw new Error("badFile"); }
    if (!d || d.format !== KEY_FORMAT || !d.salt || !d.iv || !d.data) throw new Error("badFile");
    const iter = Number.isInteger(d.iter) && d.iter >= 100000 && d.iter <= 5000000 ? d.iter : KDF_ITER;
    let plain;
    try {
      const key = await deriveKey(password, unb64(d.salt), iter);
      plain = await subtle().decrypt({ name: "AES-GCM", iv: unb64(d.iv) }, key, unb64(d.data));
    } catch { throw new Error("badPassword"); }
    let p;
    try { p = JSON.parse(new TextDecoder().decode(plain)); } catch { throw new Error("badFile"); }
    if (!p || !Array.isArray(p.changes)) throw new Error("badFile");
    return {
      file: String(p.file || ""), created: String(p.created || ""),
      changes: p.changes.filter((x) => x && typeof x.s === "string" && Number.isInteger(x.r) && Number.isInteger(x.c) && x.r >= 0 && x.c >= 0)
        .map((x) => ({ sheet: x.s, r: x.r, c: x.c, before: decVal(x.b), after: decVal(x.a) })),
    };
  }
  // Czy obecna wartość komórki to ta sama, którą wpisała anonimizacja (wtedy wolno przywrócić).
  function sameValue(a, b) {
    const ka = a == null || a === "" ? "" : a instanceof Date ? `d${a.getTime()}` : typeof a === "number" ? `n${a}` : `s${String(a)}`;
    const kb = b == null || b === "" ? "" : b instanceof Date ? `d${b.getTime()}` : typeof b === "number" ? `n${b}` : `s${String(b)}`;
    if (ka === kb) return true;
    // liczba zapisana jako tekst w pliku i odwrotnie
    return a != null && b != null && !(a instanceof Date) && !(b instanceof Date) && String(a).trim() === String(b).trim();
  }

  // ── Etap 5: AI ✦ w wolnym tekście ───────────────────────────────────────
  // Model (local-ai.js, na urządzeniu) dostaje ponumerowane teksty i zwraca fragmenty z danymi
  // osobowymi, których nie łapią reguły („rozmawiałem z panią Teresą z Lipowej”). Wynik przechodzi
  // przez sito: tylko znane rodzaje i fragment DOSŁOWNIE obecny w swoim tekście. To tylko
  // propozycje — w oknie są odznaczone, dopóki użytkownik ich nie zaznaczy.
  const FREE_HEADER_RE = /(uwag|opis|notat|koment|tresc|wiadomo|przebieg|ustalen|zgloszen|historia|remark|note|comment|descr|message|details|summary)/;
  // Czy kolumna to wolny tekst: nagłówek albo zdania (≥ 4 słowa w ≥ 30% niepustych komórek).
  function isFreeTextColumn(header, values) {
    if (FREE_HEADER_RE.test(fold(header))) return true;
    const strs = values.filter((v) => typeof v === "string" && v.trim());
    if (strs.length < 2) return false;
    const long = strs.filter((v) => v.trim().split(/\s+/).length >= 4).length;
    return long / strs.length >= 0.3;
  }
  const AI_TYPE_MAP = { osoba: "person", person: "person", nazwisko: "person", imie: "person", adres: "address", address: "address", ulica: "address", miejscowosc: "place", miasto: "place", place: "place", town: "place", city: "place" };
  const AI_SYSTEM = [
    "Jesteś narzędziem do anonimizacji. Dostajesz ponumerowane teksty z arkusza.",
    "Znajdź w nich dane osobowe: imiona i nazwiska (także samo imię albo samo nazwisko osoby, z tytułem lub bez: „pani Teresa”, „Nowakowa”), adresy (ulica z numerem lub bez) i miejscowości zamieszkania.",
    "Nie zgłaszaj nazw firm, urzędów, miesięcy, dni tygodnia, zwykłych słów ani numerów.",
    "Każdy fragment skopiuj DOKŁADNIE z tekstu, bez zmian. Nie wymyślaj.",
    'Odpowiedz wyłącznie JSON: {"znalezione": [{"nr": 1, "fragment": "…", "typ": "osoba|adres|miejscowosc"}]}. Gdy nic nie ma: {"znalezione": []}.',
  ].join(" ");
  const AI_SHOTS = [[
    "1. Dzwoniła pani Teresa z Lipowej 5, prosi o kontakt po 15.\n2. Faktura do 10 maja, przypomnieć działowi kadr.",
    '{"znalezione": [{"nr": 1, "fragment": "pani Teresa", "typ": "osoba"}, {"nr": 1, "fragment": "Lipowej 5", "typ": "adres"}]}',
  ]];
  function aiPrompt(texts, maxLen = 400) {
    return { system: AI_SYSTEM, shots: AI_SHOTS, user: texts.map((t, i) => `${i + 1}. ${String(t).replace(/\s+/g, " ").slice(0, maxLen)}`).join("\n") };
  }
  const AI_STOP = new Set("pan pani panu panem panią państwo dzień dobry poniedziałek wtorek środa czwartek piątek sobota niedziela styczeń luty marzec kwiecień maj czerwiec lipiec sierpień wrzesień październik listopad grudzień".split(" "));
  // Zwraca [{ i, type, value }] — tylko fragmenty obecne w tekście nr i (oryginalna pisownia z tekstu).
  function parseAiAnswer(texts, value) {
    const list = value && Array.isArray(value.znalezione) ? value.znalezione : value && Array.isArray(value.found) ? value.found : [];
    const out = [];
    const seen = new Set();
    for (const it of list.slice(0, 200)) {
      if (!it || typeof it !== "object") continue;
      const i = Number(it.nr ?? it.n) - 1;
      const type = AI_TYPE_MAP[fold(it.typ ?? it.type ?? "")];
      const frag = String(it.fragment ?? it.tekst ?? it.text ?? "").replace(/\s+/g, " ").trim();
      if (!Number.isInteger(i) || i < 0 || i >= texts.length || !type) continue;
      if (frag.length < 3 || frag.length > 80 || !/\p{L}/u.test(frag) || /^[\d\s.,:/-]+$/.test(frag)) continue;
      if (AI_STOP.has(norm(frag))) continue;
      const src = String(texts[i]).replace(/\s+/g, " ");
      let at = src.indexOf(frag);
      if (at < 0) at = src.toLocaleLowerCase("pl").indexOf(frag.toLocaleLowerCase("pl"));
      if (at < 0) continue; // model wymyślił albo przekręcił — odrzucamy
      const exact = src.slice(at, at + frag.length);
      // osoba musi mieć choć jedno słowo z wielkiej litery (inaczej to zwykłe słowo)
      if (type === "person" && !/(^|[\s(„"])\p{Lu}\p{Ll}/u.test(exact)) continue;
      // „pani Teresa” → osoba to „Teresa” (tytuł zostaje w tekście)
      const value2 = type === "person" ? exact.replace(/^(pan(i|u|em|ią|a)?|p\.|dr|mgr|inż\.?|prof\.?)\s+/iu, "") : exact;
      if (value2.length < 3) continue;
      const k = `${type}|${norm(value2)}`;
      if (seen.has(k)) continue;
      seen.add(k);
      out.push({ i, type, value: value2 });
    }
    return out;
  }
  // Trafienia terminów AI w tekście komórki: całe słowa, bez nakładania na istniejące trafienia.
  function aiHits(text, terms, existing = []) {
    const s = String(text ?? "");
    if (!s || !terms.length) return existing;
    const hits = existing.slice();
    const esc = (x) => x.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    for (const t of [...terms].sort((a, b) => b.value.length - a.value.length)) {
      const reT = new RegExp(`(?<![\\p{L}\\p{N}])${esc(t.value).replace(/\\? /g, "\\s+")}(?![\\p{L}\\p{N}])`, "giu");
      for (const m of s.matchAll(reT)) {
        const start = m.index, end = m.index + m[0].length;
        if (hits.some((h) => start < h.end && end > h.start)) continue;
        hits.push({ type: t.type, start, end, value: m[0], conf: "ai" });
      }
    }
    return hits.sort((a, b) => a.start - b.start);
  }

  return {
    isFreeTextColumn, aiPrompt, parseAiAnswer, aiHits,
    parseCommand, PRESETS, sanitizeProfile, exportProfiles, importProfiles, PROFILE_FORMAT,
    isQuasiColumn, riskReport, encryptReverseKey, decryptReverseKey, sameValue, KEY_FORMAT, fold,
  };
})();

if (typeof module !== "undefined" && module.exports) module.exports = SwbAnonTools;
