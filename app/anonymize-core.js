// anonymize-core.js — silnik anonimizacji (bez DOM): wykrywanie danych osobowych i metody zamiany.
//
// Zasady (2026-10-31):
//  • Wykrywanie to reguły i słowniki. Numery sprawdzane sumą kontrolną (PESEL, NIP, REGON, IBAN,
//    dowód, paszport, karta — Luhn, księga wieczysta), reszta wzorcem (e-mail, telefon, kod,
//    adres, rejestracja, VIN, IP) albo słownikiem imion i nazwisk.
//  • Kolumna liczy się jako „osoby”, „PESEL”… z nagłówka albo z tego, co w niej siedzi — wtedy
//    zamienia się CAŁĄ komórkę (także nietypowe nazwiska spoza słownika).
//  • Zamiana jest SPÓJNA w całym pliku: ten sam oryginał = ten sam wynik (pseudonim, fałszywe
//    dane, skrót). „A. Kamiński” i „Adam Kamiński” to ta sama osoba, więc dostają tę samą
//    fałszywą osobę w tej samej formie.
//  • Sesja zbiera oryginały (leakTerms) — po zapisie sprawdzamy, czy nie zostały gdzieś w pliku.
//
// API: SwbAnon.createSession(opts) → sesja; SwbAnon.findInText, profileColumn, headerHint,
//      validators, TYPES, METHODS, DEFAULT_METHODS. Ładowane też w Node (testy).

const SwbAnon = (() => {
  // ── Słowniki ────────────────────────────────────────────────────────────
  const FEMALE = "Anna Maria Katarzyna Małgorzata Agnieszka Barbara Ewa Krystyna Elżbieta Zofia Joanna Magdalena Monika Teresa Danuta Aleksandra Natalia Irena Beata Jadwiga Dorota Halina Karolina Janina Jolanta Marta Grażyna Iwona Alicja Justyna Bożena Renata Urszula Wiesława Paulina Agata Julia Hanna Sylwia Stanisława Helena Weronika Patrycja Zuzanna Wioletta Emilia Kinga Izabela Martyna Lucyna Wanda Edyta Mariola Dominika Kamila Genowefa Marianna Anita Klaudia Gabriela Daria Sabina Oliwia Milena Bogumiła Lena Ilona Wiktoria Jagoda Maja Amelia Laura Aneta Ewelina Magda Kazimiera Józefa Czesława Henryka Leokadia Regina Violetta Wioleta Roksana Adrianna Nikola Michalina Antonina Pola Malwina Marzena Bernadeta Celina Elwira Felicja Franciszka Gertruda Honorata Ida Inga Jowita Judyta Kornelia Ksenia Ludmiła Łucja Marcelina Marlena Matylda Nadia Olga Paula Sandra Stefania Tamara Zdzisława Żaneta Aniela Bronisława Lidia Liliana Diana Angelika Ada Alina Emilia Eliza Edyta Grażyna Hanna Iga Izabella Jagna Kaja Kalina Lilianna Luiza Mirosława Nina Patrycja Sara Tatiana Urszula Zenona Bogusława Dagmara Ewelina Gabriela Halina".split(" ");
  const MALE = "Piotr Krzysztof Andrzej Tomasz Jan Paweł Michał Marcin Stanisław Jakub Adam Marek Łukasz Grzegorz Mateusz Wojciech Mariusz Dariusz Zbigniew Jerzy Maciej Rafał Robert Kamil Józef Janusz Ryszard Tadeusz Jacek Kazimierz Sławomir Henryk Szymon Dawid Przemysław Daniel Damian Sebastian Artur Patryk Bartosz Roman Leszek Mirosław Karol Waldemar Zdzisław Kacper Edward Filip Hubert Czesław Bogdan Marian Wiesław Jarosław Dominik Adrian Arkadiusz Mikołaj Antoni Franciszek Igor Oskar Aleksander Wiktor Bartłomiej Konrad Norbert Radosław Sylwester Tymon Witold Władysław Zenon Leon Ignacy Julian Kajetan Ksawery Lucjan Ludwik Miłosz Natan Olaf Remigiusz Seweryn Stefan Wacław Zygmunt Bolesław Bronisław Emil Eryk Fabian Gabriel Hieronim Ireneusz Jędrzej Juliusz Kornel Krystian Leonard Lech Maksymilian Maurycy Nikodem Oliwer Kuba Tymoteusz Wincenty Albert Alan Cezary Eugeniusz Feliks Florian Gustaw Teodor Tobiasz Włodzimierz Bogusław Mieczysław Edmund Józef Lesław Michał Piotr Przemek Zbyszek Bogumił Borys Dionizy Fryderyk Gerard Hilary Ignacy Jerzy Kryspin Leopold Marcel Oktawian Patryk Rajmund Roch Stanisław Wit".split(" ");
  const SURNAMES = "Nowak Kowalski Wiśniewski Wójcik Kowalczyk Kamiński Lewandowski Zieliński Szymański Woźniak Dąbrowski Kozłowski Jankowski Mazur Wojciechowski Kwiatkowski Krawczyk Kaczmarek Piotrowski Grabowski Zając Pawłowski Michalski Król Wieczorek Jabłoński Wróbel Nowakowski Majewski Olszewski Stępień Malinowski Jaworski Adamczyk Dudek Nowicki Pawlak Górski Witkowski Walczak Sikora Baran Rutkowski Michalak Szewczyk Ostrowski Tomaszewski Pietrzak Zalewski Wróblewski Marciniak Jasiński Zawadzki Bąk Jakubowski Sadowski Duda Włodarczyk Wilk Chmielewski Borkowski Sokołowski Szczepański Sawicki Lis Kucharski Kalinowski Wysocki Mazurek Kubiak Maciejewski Czarnecki Kołodziej Urbański Kaźmierczak Sobczak Konieczny Głowacki Zakrzewski Krupa Wasilewski Krajewski Adamski Laskowski Ziółkowski Szulc Makowski Czerwiński Baranowski Szymczak Brzeziński Kozieł Kaczmarczyk Kołodziejczyk Bielak Bielecki Kurek Górecki Przybylski Kowal Leśniak Gajewski Szczepaniak Cieślak Kowalewski Kozak Ryba Chojnacki Sikorski Kasprzak Kurowski Domański Łuczak Zych Wawrzyniak Pająk Sowa Mróz Kubacki Lipiński Ciesielski Wilczek Banaś Kot Socha Rogalski Wójtowicz Kurzawa Fijałkowski Sobolewski Gołębiowski Polak Kaczor Grzelak Nowacki Sosnowski Kowalik Czaja Jędrzejewski Matuszewski Kubica Łukasik Kozieł Szewczak Rybak Kalisz Śliwiński Ostrowska Tomczak Lisowski Jankowiak Grzybowski Dziedzic Żak Bednarek Wrona Gajda Pilipczuk Maślankiewicz Bednarczyk Kamińska Urban".split(" ");
  const STREETS = "Lipowa Polna Leśna Słoneczna Krótka Szkolna Ogrodowa Łąkowa Brzozowa Kwiatowa Kościelna Sosnowa Zielona Parkowa Akacjowa Długa Kolejowa Klonowa Wiśniowa Jasna Spacerowa Cicha Topolowa Dębowa Graniczna Wierzbowa Jesionowa Bukowa Modrzewiowa Różana".split(" ");
  const CITIES = "Kraków Warszawa Poznań Gdańsk Wrocław Łódź Lublin Szczecin Toruń Kielce Opole Rzeszów Olsztyn Białystok Zielona Góra Gorzów Bydgoszcz Katowice Radom Płock".split(" ");

  const norm = (s) => String(s ?? "").normalize("NFC").trim().replace(/\s+/g, " ").toLocaleLowerCase("pl");
  const fold = (s) => norm(s).normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/ł/g, "l");
  const FIRST = new Map();
  FEMALE.forEach((n) => FIRST.set(norm(n), "f"));
  MALE.forEach((n) => FIRST.set(norm(n), "m"));
  const SURNAME_SET = new Set(SURNAMES.map(norm));
  const FEM_FIRST = [...new Set(FEMALE)];
  const MALE_FIRST = [...new Set(MALE)];
  const FAKE_SURNAMES = [...new Set(SURNAMES.filter((s) => !/(ska|cka|dzka)$/.test(s)))];
  // nazwisko polskie z końcówki — tylko jako uzupełnienie (w kolumnie osób albo przy imieniu)
  const SURNAME_SUFFIX = /(ski|ska|cki|cka|dzki|dzka|wicz|owicz|ewicz|czyk|czak|czuk|szczak|ańczyk|ak|ek|ik|yk|uk|iak|iec|ło|la|ny|na)$/i;
  const SURNAME_SUFFIX_STRONG = /(ski|ska|cki|cka|dzki|dzka|wicz|czyk|czuk|ńczyk)$/i;
  const NOT_NAMES = new Set(["ul", "al", "os", "pl", "nr", "tel", "pan", "pani", "dr", "mgr", "inż", "suma", "razem", "uwagi", "brak", "tak", "nie", "status", "data", "kwota", "opis", "zamknięte", "w trakcie"].map(norm));

  // ── Walidatory ──────────────────────────────────────────────────────────
  const digits = (s) => String(s).replace(/\D/g, "");
  function peselDate(p) {
    let y = +p.slice(0, 2), m = +p.slice(2, 4);
    const d = +p.slice(4, 6);
    let century = 1900;
    if (m > 80) { century = 1800; m -= 80; } else if (m > 60) { century = 2200; m -= 60; } else if (m > 40) { century = 2100; m -= 40; } else if (m > 20) { century = 2000; m -= 20; }
    y += century;
    const dt = new Date(Date.UTC(y, m - 1, d));
    if (m < 1 || m > 12 || dt.getUTCDate() !== d) return null;
    return { y, m, d };
  }
  function peselCheck(p10) {
    const w = [1, 3, 7, 9, 1, 3, 7, 9, 1, 3];
    let s = 0;
    for (let i = 0; i < 10; i++) s += (+p10[i]) * w[i];
    return (10 - (s % 10)) % 10;
  }
  function peselValid(s) {
    const p = String(s);
    if (!/^\d{11}$/.test(p)) return false;
    if (peselCheck(p) !== +p[10]) return false;
    return !!peselDate(p);
  }
  function nipCheck(n9) {
    const w = [6, 5, 7, 2, 3, 4, 5, 6, 7];
    let s = 0;
    for (let i = 0; i < 9; i++) s += (+n9[i]) * w[i];
    return s % 11;
  }
  function nipValid(s) {
    const n = digits(s);
    if (n.length !== 10 || /^0{3}/.test(n)) return false;
    const c = nipCheck(n);
    return c !== 10 && c === +n[9];
  }
  function regonCheck(body) {
    const w = body.length === 8 ? [8, 9, 2, 3, 4, 5, 6, 7] : [2, 4, 8, 5, 0, 9, 7, 3, 6, 1, 2, 4, 8];
    let s = 0;
    for (let i = 0; i < w.length; i++) s += (+body[i]) * w[i];
    const c = s % 11;
    return c === 10 ? 0 : c;
  }
  function regonValid(s) {
    const n = digits(s);
    if (n.length !== 9 && n.length !== 14) return false;
    if (/^0+$/.test(n)) return false;
    return regonCheck(n.slice(0, -1)) === +n[n.length - 1];
  }
  function mod97(str) {
    let r = 0;
    for (const ch of str) r = (r * 10 + (+ch)) % 97;
    return r;
  }
  function ibanNumeric(iban) {
    const s = iban.slice(4) + iban.slice(0, 4);
    return s.replace(/[A-Z]/g, (c) => String(c.charCodeAt(0) - 55));
  }
  function ibanValid(s) {
    let v = String(s).toUpperCase().replace(/[\s-]/g, "");
    if (/^\d{26}$/.test(v)) v = "PL" + v; // NRB bez kraju
    if (!/^[A-Z]{2}\d{2}[A-Z0-9]{10,30}$/.test(v)) return false;
    return mod97(ibanNumeric(v)) === 1;
  }
  function luhnValid(s) {
    const n = digits(s);
    if (n.length < 13 || n.length > 19) return false;
    let sum = 0;
    for (let i = 0; i < n.length; i++) {
      let d = +n[n.length - 1 - i];
      if (i % 2 === 1) { d *= 2; if (d > 9) d -= 9; }
      sum += d;
    }
    return sum % 10 === 0;
  }
  const letterVal = (c) => (/\d/.test(c) ? +c : c.charCodeAt(0) - 55);
  function idCardCheck(s) { // ABC 1 23456 — 4. znak = cyfra kontrolna
    const w = [7, 3, 1, 0, 7, 3, 1, 7, 3];
    let sum = 0;
    for (let i = 0; i < 9; i++) if (i !== 3) sum += letterVal(s[i]) * w[i];
    return sum % 10;
  }
  function idCardValid(s) {
    const v = String(s).toUpperCase().replace(/\s/g, "");
    return /^[A-Z]{3}\d{6}$/.test(v) && idCardCheck(v) === +v[3];
  }
  function passportCheck(s) { // AB 1 234567 — 3. znak = cyfra kontrolna
    const w = [7, 3, 0, 1, 7, 3, 1, 7, 3];
    let sum = 0;
    for (let i = 0; i < 9; i++) if (i !== 2) sum += letterVal(s[i]) * w[i];
    return sum % 10;
  }
  function passportValid(s) {
    const v = String(s).toUpperCase().replace(/\s/g, "");
    return /^[A-Z]{2}\d{7}$/.test(v) && passportCheck(v) === +v[2];
  }
  const KW_MAP = "0123456789X.ABCDEFGHIJKLMNOPRSTUWYZ"; // X=10, A=11 … Z=33 (bez Q, V)
  function kwVal(c) {
    if (/\d/.test(c)) return +c;
    if (c === "X") return 10;
    const i = "ABCDEFGHIJKLMNOPRSTUWYZ".indexOf(c);
    return i < 0 ? 0 : 11 + i;
  }
  function kwCheck(court, num) {
    const s = court + num;
    const w = [1, 3, 7];
    let sum = 0;
    for (let i = 0; i < s.length; i++) sum += kwVal(s[i]) * w[i % 3];
    return sum % 10;
  }
  function kwValid(s) {
    const m = String(s).toUpperCase().match(/^([A-Z]{2}\d[A-Z])\/(\d{8})\/(\d)$/);
    return !!m && kwCheck(m[1], m[2]) === +m[3];
  }
  void KW_MAP;

  // ── Typy i metody ───────────────────────────────────────────────────────
  // Kolejność = pierwszeństwo przy nakładaniu się trafień.
  const TYPES = ["custom", "person", "pesel", "nip", "regon", "iban", "card", "idcard", "passport", "kw", "email", "phone", "address", "postcode", "place", "dob", "plate", "vin", "ip", "other"];
  const METHODS = ["keep", "pseudo", "fake", "mask", "generalize", "redact", "remove", "hash", "shift"];
  const DEFAULT_METHODS = {
    custom: "redact", person: "pseudo", pesel: "mask", nip: "mask", regon: "mask", iban: "mask", card: "mask",
    idcard: "mask", passport: "mask", kw: "mask", email: "mask", phone: "mask", address: "pseudo",
    postcode: "generalize", place: "keep", dob: "generalize", plate: "mask", vin: "mask", ip: "mask", other: "redact",
  };

  // ── Adresy ──────────────────────────────────────────────────────────────
  // „ul. Długa 5/3”, „al. Jana Pawła II 12 m. 4”, „ul. 3 Maja 1”, „ul. gen. Andersa 7, 00-950 Warszawa”,
  // „os. Słoneczne 4”, „ulica Długa” (bez numeru — średnia pewność), „Długa 5, 00-950 Warszawa”
  // (bez prefiksu, ale z kodem). Sama „Długa 5” bez prefiksu i kodu = tylko w kolumnie adresów
  // (nagłówek „Adres”, „Ulica” albo treść kolumny — addressCellScore), bo „Grupa 3”, „Rozdział 5”.
  const ci = (w) => w.replace(/\p{L}/gu, (ch) => `[${ch.toLowerCase()}${ch.toUpperCase()}]`);
  const STREET_PREFIXES = ["ulica", "ul\\.?", "aleje", "aleja", "al\\.", "osiedle", "os\\.", "plac", "pl\\.", "rondo", "skwer", "bulwar", "bulw\\.", "rynek", "trakt"];
  const ADDR_PREFIX = `(?:${STREET_PREFIXES.map(ci).join("|")})`;
  const ADDR_TOK = "(?:\\p{Lu}[\\p{L}'’-]*\\.?|\\p{Ll}{2,6}\\.)"; // „Długa”, „II”, „gen.”, „św.”
  const ADDR_NAME = `(?:${ADDR_TOK}|\\d{1,2})(?:\\s+${ADDR_TOK}){0,4}`; // „3 Maja”
  const ADDR_HOUSE = "\\d{1,4}[a-zA-Z]?(?:\\s?[/-]\\s?\\d{1,4}[a-zA-Z]?)?(?:,?\\s*(?:m\\.?|lok\\.?|lokal|mieszk\\.?)\\s*\\d{1,4}[a-zA-Z]?)?";
  const ADDR_TOWN = "\\p{Lu}\\p{Ll}+(?:[ -]\\p{Lu}\\p{Ll}+)?";
  const ADDR_TAIL = `(?:,?\\s*\\d{2}-\\d{3}(?:\\s+${ADDR_TOWN})?)?`;
  const ADDR_RE_PREFIXED = new RegExp(`(?<![\\p{L}\\p{N}])${ADDR_PREFIX}\\s*${ADDR_NAME}\\s+${ADDR_HOUSE}${ADDR_TAIL}(?![\\p{L}\\p{N}])`, "gu");
  const ADDR_RE_STREET_ONLY = new RegExp(`(?<![\\p{L}\\p{N}])${ADDR_PREFIX}\\s*${ADDR_NAME}(?![\\p{L}\\p{N}])(?!\\s*\\d)`, "gu");
  const ADDR_NAME_PLAIN = "\\p{Lu}[\\p{L}'’-]*\\.?(?:\\s+(?:\\p{Lu}[\\p{L}'’-]*\\.?|\\p{Ll}{2,6}\\.)){0,3}";
  const ADDR_RE_POSTCODE = new RegExp(`(?<![\\p{L}\\p{N}])(?:${ADDR_NAME_PLAIN}|\\d{1,2}\\s+${ADDR_NAME_PLAIN})\\s+${ADDR_HOUSE},?\\s*\\d{2}-\\d{3}(?:\\s+${ADDR_TOWN})?(?![\\p{L}\\p{N}])`, "gu");
  const ADDR_RE_CELL = new RegExp(`^(?:${ADDR_PREFIX}\\s*)?(${ADDR_NAME})\\s+(\\d{1,4})[a-zA-Z]?(?:\\s?[/-]\\s?\\d{1,4}[a-zA-Z]?)?(?:,?\\s*(?:m\\.?|lok\\.?|lokal|mieszk\\.?)\\s*\\d{1,4}[a-zA-Z]?)?${ADDR_TAIL}$`, "u");
  const RE_ADDR_PREFIX_START = new RegExp(`^${ADDR_PREFIX}\\s*`, "u");

  // Cała komórka wygląda na adres (dla profilu kolumny): 1 = z prefiksem/kodem, 0.7 = „Długa 5”.
  // Numer z zerami wiodącymi („Osoba 010”) to nie numer domu.
  function addressCellScore(v) {
    const s = String(v ?? "").trim();
    if (!s || s.length > 90) return 0;
    const m = s.match(ADDR_RE_CELL);
    if (!m || /^0\d/.test(m[2])) return 0;
    return RE_ADDR_PREFIX_START.test(s) || /\d{2}-\d{3}/.test(s) ? 1 : 0.7;
  }
  // Nazwa ulicy z komórki-adresu (bez prefiksu i numeru) — do oceny różnorodności kolumny.
  function addressStreet(v) {
    const m = String(v ?? "").trim().match(ADDR_RE_CELL);
    return m ? norm(m[1]) : "";
  }

  // Wzorce w tekście. validate(m) → "high" | "medium" | null.
  const B = "(?<![\\p{L}\\p{N}])"; // granica „nie litera/cyfra” z przodu (Unicode)
  const E = "(?![\\p{L}\\p{N}])";
  const rx = (src) => new RegExp(src, "gu");
  const DETECTORS = [
    { type: "iban", re: rx(`${B}(?:PL[ ]?)?\\d{2}(?:[ -]?\\d{4}){6}${E}`), validate: (s) => (ibanValid(s) ? "high" : null) },
    { type: "iban", re: rx(`${B}[A-Z]{2}\\d{2}(?:[ ]?[A-Z0-9]{4}){2,7}(?:[ ]?[A-Z0-9]{1,4})?${E}`), validate: (s) => (/^PL/i.test(s) ? null : ibanValid(s) ? "high" : null) },
    { type: "card", re: rx(`${B}\\d{4}(?:[ -]?\\d{4}){2}[ -]?\\d{1,7}${E}`), validate: (s) => (luhnValid(s) && /[ -]/.test(s) ? "high" : luhnValid(s) && digits(s).length === 16 ? "medium" : null) },
    { type: "pesel", re: rx(`${B}\\d{11}${E}`), validate: (s) => (peselValid(s) ? "high" : null) },
    { type: "nip", re: rx(`${B}(?:PL[ ]?)?(?:\\d{3}-\\d{3}-\\d{2}-\\d{2}|\\d{3}-\\d{2}-\\d{2}-\\d{3}|\\d{3} \\d{3} \\d{2} \\d{2}|\\d{10})${E}`), validate: (s, ctx) => (nipValid(s) ? (/[- ]/.test(s.replace(/^PL ?/i, "")) || /nip/i.test(ctx) ? "high" : "medium") : null) },
    { type: "regon", re: rx(`${B}(?:\\d{9}|\\d{14})${E}`), validate: (s, ctx) => (regonValid(s) && /regon/i.test(ctx) ? "high" : null) },
    { type: "kw", re: rx(`${B}[A-Z]{2}\\d[A-Z]\\/\\d{8}\\/\\d${E}`), validate: (s) => (kwValid(s) ? "high" : "medium") },
    { type: "idcard", re: rx(`${B}[A-Z]{3}[ ]?\\d{6}${E}`), validate: (s) => (idCardValid(s) ? "high" : null) },
    { type: "passport", re: rx(`${B}[A-Z]{2}[ ]?\\d{7}${E}`), validate: (s) => (passportValid(s) ? "high" : null) },
    { type: "email", re: /(?<![\w.%+-])[\w.%+-]+@[\p{L}\p{N}-]+(?:\.[\p{L}\p{N}-]+)*\.[\p{L}]{2,}(?![\w-])/gu, validate: () => "high" },
    { type: "ip", re: rx(`${B}(?:(?:25[0-5]|2[0-4]\\d|1?\\d?\\d)\\.){3}(?:25[0-5]|2[0-4]\\d|1?\\d?\\d)${E}`), validate: (s, ctx) => (/\bip\b|adres ip|serwer|host/i.test(ctx) || !/^\d{1,2}\.\d{1,2}\.\d{1,2}\.\d{1,2}$/.test(s) ? "medium" : null) },
    { type: "phone", re: rx(`${B}(?:(?:\\+|00)48[ -]?)?(?:\\d{3}[ -]\\d{3}[ -]\\d{3}|\\d{2}[ -]\\d{3}[ -]\\d{2}[ -]\\d{2}|\\(\\d{2}\\)[ -]?\\d{3}[ -]?\\d{2}[ -]?\\d{2})${E}`), validate: (s) => (digits(s).replace(/^(00)?48/, "").length === 9 ? "high" : null) },
    { type: "phone", re: rx(`${B}(?:(?:\\+|00)48[ -]?)?\\d{9}${E}`), validate: (s, ctx) => (/^(\+|00)48/.test(s) || /tel|kom|phone|nr\s*tel|dzwo/i.test(ctx) ? "high" : null) },
    // adresy: z prefiksem (z numerem i bez), bez prefiksu tylko z kodem pocztowym — patrz ADDR_* niżej
    { type: "address", re: ADDR_RE_PREFIXED, validate: () => "high" },
    { type: "address", re: ADDR_RE_POSTCODE, validate: () => "high" },
    { type: "address", re: ADDR_RE_STREET_ONLY, validate: () => "medium" },
    { type: "postcode", re: rx(`${B}\\d{2}-\\d{3}${E}`), validate: (s, ctx) => (/^\d{2}-\d{3}\s+\p{Lu}/u.test(ctx.slice(ctx.indexOf(s))) || /kod|poczt/i.test(ctx) ? "high" : "medium") },
    { type: "plate", re: rx(`${B}[BCDEFGKLNOPRSTWZ][A-Z]{1,2} (?=[0-9A-Z]*\\d)[0-9A-Z]{4,5}${E}`), validate: () => "medium" },
    { type: "vin", re: rx(`${B}(?=[A-HJ-NPR-Z0-9]*\\d[A-HJ-NPR-Z0-9]*\\d)(?=[A-HJ-NPR-Z0-9]*[A-HJ-NPR-Z][A-HJ-NPR-Z0-9]*[A-HJ-NPR-Z])[A-HJ-NPR-Z0-9]{17}${E}`), validate: () => "medium" },
  ];

  // Osoba w wolnym tekście: imię ze słownika + słowo z wielkiej litery (w dowolnej kolejności),
  // „Pan/Pani/dr …” + nazwisko, inicjał + nazwisko o mocnej końcówce (-ski, -wicz…) lub ze słownika.
  const NAME_WORD = "\\p{Lu}[\\p{Ll}]+(?:-\\p{Lu}[\\p{Ll}]+)?";
  const RE_INITIAL = new RegExp(`(?<![\\p{L}\\p{N}])(\\p{Lu})(?:\\.\\s?|\\s)(${NAME_WORD})(?![\\p{L}])`, "gu");
  const RE_TITLE = new RegExp(`(?<![\\p{L}])(?:[Pp]an(?:i|a|u|em|ią)?|[Dd]r|[Mm]gr|[Ii]nż\\.|[Pp]rof\\.)\\s+(${NAME_WORD}(?:\\s+${NAME_WORD})?)(?![\\p{L}])`, "gu");

  // Imiona i nazwiska poznane z kolumn osób w pliku (spoza słownika: „Faustyna”, „Żmuda”).
  // Ustawia je sesja (activate) — każda sesja ma własne, plik nie wpływa na inny plik.
  let LEARNED = { first: new Map(), sur: new Set() };
  const isFirst = (w) => { const n = norm(w); return FIRST.has(n) || LEARNED.first.has(n); };
  const firstGender = (w) => { const n = norm(w); return FIRST.get(n) || LEARNED.first.get(n) || null; };
  const knownSurname = (w) => { const n = norm(w); return SURNAME_SET.has(n) || LEARNED.sur.has(n); };
  const surnameLike = (w, strong) => {
    const n = norm(w);
    if (NOT_NAMES.has(n) || isFirst(n)) return false;
    if (knownSurname(n)) return true;
    return n.split("-").every((part) => (strong ? SURNAME_SUFFIX_STRONG : SURNAME_SUFFIX).test(part) && part.length >= 3);
  };

  // Imię w dowolnym przypadku: andrzeja→andrzej, pawła→paweł (wstawne e), anny/annie/annę/anną→anna.
  const NOT_FIRST_INFL = new Set(["maja", "maju", "róży"]); // „3 maja”
  function firstInfl(w) {
    const n = norm(w);
    if (NOT_FIRST_INFL.has(n)) return false;
    if (isFirst(n)) return true;
    const cands = [];
    for (const suf of ["owi", "iem", "em", "ie", "a", "u", "y", "i", "ę", "ą", "o"]) {
      if (!n.endsWith(suf) || n.length - suf.length < 2) continue;
      const stem = n.slice(0, -suf.length);
      cands.push(stem, `${stem}a`, stem.replace(/(\p{L})$/u, "e$1"));
      if (suf === "ie") cands.push(stem.replace(/n$/, "na"), `${stem}a`);
    }
    return cands.some((c) => FIRST.has(c) || LEARNED.first.has(c));
  }
  // Nazwisko w odmianie: -skiego/-skiemu/-skim/-ską/-skiej, -wicza/-wiczem, -czyka…, albo poznane z pliku.
  const SURNAME_INFL = /(?:sk|ck|dzk)(?:i|a|iego|iemu|im|ą|iej|ich)$|wicz(?:a|owi|em|u|ów)?$|czyk(?:a|owi|iem|u)?$|czuk(?:a|owi|iem|u)?$/i;
  function surnameInfl(w) {
    const n = norm(w);
    if (n.length < 4 || NOT_NAMES.has(n)) return false;
    if (SURNAME_INFL.test(n) || knownSurname(n)) return true;
    for (const suf of ["owi", "iem", "em", "a", "u", "ą", "ej"]) if (n.endsWith(suf) && knownSurname(n.slice(0, -suf.length))) return true;
    return false;
  }

  // Ciąg słów z wielkiej litery; osoba = imię ze słownika + sąsiednie nazwisko (najpierw
  // następne słowo — „Jan Kowalski”, potem poprzednie — „Kowalski Jan”). Słowo przed imieniem
  // na początku zdania („Dzwonił Jan Kowalski”) nie wchodzi do trafienia.
  const RE_CAPWORD = new RegExp(`(?<![\\p{L}.])${NAME_WORD}(?![\\p{L}])`, "gu");
  function findPersons(text, out) {
    const words = [...text.matchAll(RE_CAPWORD)].map((m) => ({ w: m[0], start: m.index, end: m.index + m[0].length }));
    const adjacent = (a, b) => a && b && /^\s+$/.test(text.slice(a.end, b.start));
    const used = new Set();
    words.forEach((wd, i) => {
      if (used.has(i) || !isFirst(wd.w)) return;
      let from = i, to = i;
      const next = words[i + 1], prev = words[i - 1];
      if (adjacent(wd, next) && !isFirst(next.w) && !NOT_NAMES.has(norm(next.w))) to = i + 1;
      else if (adjacent(prev, wd) && !used.has(i - 1) && !isFirst(prev.w) && surnameLike(prev.w, false)) from = i - 1;
      else return;
      // drugie imię: „Jan Maria Kowalska”
      if (to === i && adjacent(words[to], words[to + 1]) && !isFirst(words[to + 1].w) && surnameLike(words[to + 1].w, false)) to++;
      if (to === i + 1 && isFirst(words[to].w) && adjacent(words[to], words[to + 1]) && surnameLike(words[to + 1].w, false)) to++;
      for (let k = from; k <= to; k++) used.add(k);
      const sur = words.slice(from, to + 1).filter((x) => !isFirst(x.w));
      const conf = sur.some((x) => surnameLike(x.w, false)) ? "high" : "medium";
      out.push({ type: "person", start: words[from].start, end: words[to].end, value: text.slice(words[from].start, words[to].end), conf });
    });
    for (const m of text.matchAll(RE_INITIAL)) {
      const dot = m[0].includes(".");
      const ok = isFirst(m[2]) ? dot : dot ? surnameLike(m[2], true) || knownSurname(m[2]) : LEARNED.sur.has(norm(m[2]));
      if (!ok) continue;
      out.push({ type: "person", start: m.index, end: m.index + m[0].length, value: m[0], conf: "medium" });
    }
    // Odmienione i pisane małą literą: „dodać andrzeja słowikowskiego”, „u anny kowalskiej”,
    // „z pawłem nowickim”. Para słów: imię (po odjęciu końcówki przypadka) + nazwisko z mocną
    // końcówką w odmianie albo nazwisko poznane z pliku. Pojedyncze imię — nie (za dużo zwykłych słów).
    const words2 = [...text.matchAll(/(?<![\p{L}])\p{L}{3,}(?:-\p{L}{3,})?(?![\p{L}])/gu)].map((m) => ({ w: m[0], start: m.index, end: m.index + m[0].length }));
    for (let i = 0; i + 1 < words2.length; i++) {
      const a = words2[i], b = words2[i + 1];
      if (!/^\s+$/.test(text.slice(a.end, b.start))) continue;
      if (/^\p{Lu}/u.test(a.w) && /^\p{Lu}/u.test(b.w) && isFirst(a.w)) continue; // zwykła forma — wyżej
      const ok = (firstInfl(a.w) && surnameInfl(b.w)) || (surnameInfl(a.w) && firstInfl(b.w));
      if (!ok) continue;
      out.push({ type: "person", start: a.start, end: b.end, value: text.slice(a.start, b.end), conf: "medium" });
      i++;
    }
    for (const m of text.matchAll(RE_TITLE)) {
      const start = m.index + m[0].indexOf(m[1]);
      out.push({ type: "person", start, end: start + m[1].length, value: m[1], conf: "medium" });
    }
  }

  function escapeRe(s) { return String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); }

  // Reguły użytkownika: terms (słowa/frazy zawsze do ukrycia), patterns (wyrażenia regularne),
  // exceptions (nigdy nie ruszać — porównanie bez wielkości liter).
  function compileRules(rules = {}) {
    const terms = (rules.terms || []).map((s) => String(s).trim()).filter((s) => s.length >= 2);
    const termRe = terms.length
      ? new RegExp(`(?<![\\p{L}\\p{N}])(?:${terms.sort((a, b) => b.length - a.length).map(escapeRe).join("|")})(?![\\p{L}\\p{N}])`, "giu")
      : null;
    const patterns = [];
    (rules.patterns || []).forEach((p) => {
      const src = String(p).trim();
      if (!src) return;
      const m = src.match(/^\/(.+)\/([a-z]*)$/);
      try { patterns.push(new RegExp(m ? m[1] : src, `${(m ? m[2] : "i").replace(/g/g, "")}gu`)); } catch { /* zły wzorzec — pomiń */ }
    });
    const exceptions = new Set((rules.exceptions || []).map(norm).filter(Boolean));
    return { termRe, patterns, exceptions, types: rules.types || null };
  }

  // Trafienia w tekście (bez nakładania się), w kolejności występowania.
  function findInText(text, compiled, ctxHint = "") {
    const s = String(text ?? "");
    if (!s) return [];
    const c = compiled || compileRules();
    const raw = [];
    if (c.termRe) for (const m of s.matchAll(c.termRe)) raw.push({ type: "custom", start: m.index, end: m.index + m[0].length, value: m[0], conf: "high" });
    c.patterns.forEach((re) => {
      re.lastIndex = 0;
      for (const m of s.matchAll(re)) if (m[0]) raw.push({ type: "custom", start: m.index, end: m.index + m[0].length, value: m[0], conf: "high" });
    });
    const ctx = `${ctxHint} ${s}`;
    for (const d of DETECTORS) {
      if (c.types && !c.types[d.type]) continue;
      d.re.lastIndex = 0;
      for (const m of s.matchAll(d.re)) {
        const conf = d.validate(m[0], ctx);
        if (conf) raw.push({ type: d.type, start: m.index, end: m.index + m[0].length, value: m[0], conf });
      }
    }
    if (!c.types || c.types.person) findPersons(s, raw);
    // „ul. Adama Mickiewicza 5” — osoba w nazwie ulicy należy do adresu (adres wygrywa)
    const addrSpans = raw.filter((h) => h.type === "address");
    if (addrSpans.length) for (let i = raw.length - 1; i >= 0; i--) {
      const h = raw[i];
      if (h.type === "person" && addrSpans.some((a) => h.start >= a.start && h.end <= a.end)) raw.splice(i, 1);
    }
    // Pierwszeństwo: typ wg TYPES, potem dłuższe trafienie.
    const rank = (h) => TYPES.indexOf(h.type);
    raw.sort((a, b) => rank(a) - rank(b) || (b.end - b.start) - (a.end - a.start) || a.start - b.start);
    const taken = [];
    const out = [];
    for (const h of raw) {
      if (c.exceptions.has(norm(h.value))) continue;
      if (taken.some(([a, b]) => h.start < b && h.end > a)) continue;
      taken.push([h.start, h.end]);
      out.push(h);
    }
    return out.sort((a, b) => a.start - b.start);
  }

  // ── Nagłówki kolumn ─────────────────────────────────────────────────────
  const HEADER_HINTS = [
    ["pesel", /\bpesel\b/],
    ["nip", /\bnip\b/],
    ["regon", /\bregon\b/],
    ["iban", /\biban\b|nr\.? ?(konta|rachunku)|numer (konta|rachunku)|rachunek bankowy|konto bankowe|\bkonto\b|\bnrb\b/],
    ["card", /karty? platnicz|nr\.? ?karty|numer karty|card number/],
    ["idcard", /dowod|nr\.? ?d\.?o\.?\b|seria i numer|id card/],
    ["passport", /paszport|passport/],
    ["kw", /ksieg[ai] wieczyst|\bkw\b|nr\.? ?kw/],
    ["email", /e-?mail|adres mailowy|\bmail\b/],
    ["phone", /telefon|\btel\b|\btel\.|komork|\bgsm\b|phone|mobile|nr\.? ?kontaktowy/],
    ["dob", /data urodzenia|\bur\.|urodzon|birth|\bdob\b/],
    ["postcode", /kod pocztowy|kod poczt|postal|zip/],
    ["address", /\badres\b|ulica|\bul\.|zamieszkani|zameldowani|address|street/],
    ["place", /miejscowosc|miasto|gmina|powiat|\bcity\b|town|miejsce urodzenia/],
    ["plate", /rejestracj|nr\.? ?rej|tablic|registration|license plate/],
    ["vin", /\bvin\b|nr\.? ?nadwozia|numer nadwozia/],
    ["ip", /\bip\b|adres ip/],
    ["person", /imie i nazwisko|imiona i nazwiska|nazwiska? i imion|nazwisk|\bimie\b|\bimiona\b|osoba|osob\b|pracownik|wlasciciel|klient|pacjent|uczen|student|kierowca|najemca|wynajmujac|odbiorca|nadawca|zglaszajac|opiekun|kierownik|przedstawiciel|pelnomocnik|wnioskodawca|inwestor|kontakt|podpis|osoba odpowiedzialna|wykonawca|zleceniodawca|\bname\b|surname|employee|owner|customer|contact person/],
  ];
  function headerHint(header) {
    const h = fold(header);
    if (!h) return null;
    for (const [type, re] of HEADER_HINTS) if (re.test(h)) return type;
    return null;
  }

  // Ocena „to wygląda na osobę” dla całej komórki (kolumna osób).
  function personCellScore(v) {
    const s = String(v ?? "").trim();
    if (!s || s.length > 60 || /\d/.test(s)) return 0;
    const words = s.replace(/[.,]/g, (m) => (m === "." ? ". " : " ")).split(/\s+/).filter(Boolean);
    if (words.length > 4) return 0;
    const cap = (w) => /^\p{Lu}/u.test(w);
    if (!words.every((w) => cap(w) || /^\p{Lu}\.$/u.test(w))) return 0;
    const firsts = words.filter(isFirst).length;
    const initials = words.filter((w) => /^\p{Lu}\.$/u.test(w)).length;
    const surn = words.filter((w) => !isFirst(w) && !/^\p{Lu}\.$/u.test(w) && surnameLike(w, false)).length;
    if (firsts && (surn || words.length >= 2)) return 1;
    if (initials && surn) return 0.9;
    if (words.length === 1 && (firsts || knownSurname(words[0]))) return 0.8; // także nazwiska poznane z pliku
    if (words.length === 1 && surnameLike(words[0], true)) return 0.6;
    // słaba końcówka (-ny, -ak, -ek…) tylko przy co najmniej dwóch słowach: „Mikrofony” to nie nazwisko
    if (surn && words.length >= 2 && words.length <= 3) return 0.5;
    return 0;
  }

  // Profil kolumny: values = [{v, isNum, isDate}], zwraca typ całej kolumny (albo null) i liczniki.
  function profileColumn(header, values, compiled) {
    const hint = headerHint(header);
    let nonEmpty = 0, personish = 0, addressish = 0;
    const streets = new Set();
    const whole = {}; // typ → ile komórek w całości jest tym typem
    for (const it of values) {
      const v = it.v;
      if (v == null || v === "") continue;
      nonEmpty++;
      if (typeof v === "string") {
        if (personCellScore(v) >= 0.8) personish++;
        if (addressCellScore(v) >= 0.7) { addressish++; streets.add(addressStreet(v)); }
        const hits = findInText(v, compiled, header);
        if (hits.length === 1 && hits[0].start === 0 && hits[0].end === v.trim().length) whole[hits[0].type] = (whole[hits[0].type] || 0) + 1;
      } else if (typeof v === "number" && Number.isInteger(v)) {
        const s = String(v);
        if (peselValid(s)) whole.pesel = (whole.pesel || 0) + 1;
        else if (s.length === 10 && nipValid(s)) whole.nip = (whole.nip || 0) + 1;
        else if (s.length === 9 && /^[4-8]/.test(s)) whole.phone9 = (whole.phone9 || 0) + 1;
      }
    }
    let type = null;
    let share = 0;
    if (hint && hint !== "person") {
      type = hint;
      share = 1;
    } else if (hint === "person") {
      type = "person";
      share = nonEmpty ? Math.max(personish / nonEmpty, 0.5) : 0;
    }
    if (!type && nonEmpty >= 3) {
      if (personish / nonEmpty >= 0.6) { type = "person"; share = personish / nonEmpty; }
      for (const [t, n] of Object.entries(whole)) {
        if (t === "phone9") continue;
        const sh = n / nonEmpty;
        if (sh >= 0.6 && sh > share) { type = t; share = sh; }
      }
      // „Długa 5”, „Lipowa 12/3” bez prefiksu: kolumna adresów, jeśli większość komórek tak wygląda
      // i ulice są różne („Grupa 1”, „Grupa 2”… to etykiety, nie adresy)
      const sh = addressish / nonEmpty;
      if (sh >= 0.6 && sh > share && personish / nonEmpty < 0.5 && streets.size >= Math.max(2, addressish * 0.4)) { type = "address"; share = sh; }
    }
    return { hint, type, share, nonEmpty };
  }

  // ── Narzędzia losowości (deterministyczne — ten sam oryginał = ten sam wynik) ──
  function fnv(str, seed = 0) {
    let h = 0x811c9dc5 ^ seed;
    for (let i = 0; i < str.length; i++) { h ^= str.charCodeAt(i); h = Math.imul(h, 0x01000193); }
    return h >>> 0;
  }
  function rng(seedStr, seed) {
    let a = fnv(seedStr, seed) || 1;
    return () => { // mulberry32
      a |= 0; a = (a + 0x6D2B79F5) | 0;
      let t = Math.imul(a ^ (a >>> 15), 1 | a);
      t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
  }

  // SHA-256 (synchronicznie, do metody „skrót”).
  function sha256Hex(msg) {
    const K = [0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5, 0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174, 0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da, 0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967, 0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85, 0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070, 0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3, 0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2];
    const bytes = Array.from(unescape(encodeURIComponent(msg)), (c) => c.charCodeAt(0));
    const bitLen = bytes.length * 8;
    bytes.push(0x80);
    while (bytes.length % 64 !== 56) bytes.push(0);
    for (let i = 7; i >= 0; i--) bytes.push(i >= 4 ? 0 : (bitLen >>> (i * 8)) & 0xff);
    const H = [0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19];
    const W = new Array(64);
    const rot = (x, n) => (x >>> n) | (x << (32 - n));
    for (let o = 0; o < bytes.length; o += 64) {
      for (let i = 0; i < 16; i++) W[i] = (bytes[o + i * 4] << 24) | (bytes[o + i * 4 + 1] << 16) | (bytes[o + i * 4 + 2] << 8) | bytes[o + i * 4 + 3];
      for (let i = 16; i < 64; i++) {
        const s0 = rot(W[i - 15], 7) ^ rot(W[i - 15], 18) ^ (W[i - 15] >>> 3);
        const s1 = rot(W[i - 2], 17) ^ rot(W[i - 2], 19) ^ (W[i - 2] >>> 10);
        W[i] = (W[i - 16] + s0 + W[i - 7] + s1) | 0;
      }
      let [a, b, c, d, e, f, g, h] = H;
      for (let i = 0; i < 64; i++) {
        const S1 = rot(e, 6) ^ rot(e, 11) ^ rot(e, 25);
        const t1 = (h + S1 + ((e & f) ^ (~e & g)) + K[i] + W[i]) | 0;
        const S0 = rot(a, 2) ^ rot(a, 13) ^ rot(a, 22);
        const t2 = (S0 + ((a & b) ^ (a & c) ^ (b & c))) | 0;
        h = g; g = f; f = e; e = (d + t1) | 0; d = c; c = b; b = a; a = (t1 + t2) | 0;
      }
      H[0] = (H[0] + a) | 0; H[1] = (H[1] + b) | 0; H[2] = (H[2] + c) | 0; H[3] = (H[3] + d) | 0;
      H[4] = (H[4] + e) | 0; H[5] = (H[5] + f) | 0; H[6] = (H[6] + g) | 0; H[7] = (H[7] + h) | 0;
    }
    return H.map((x) => (x >>> 0).toString(16).padStart(8, "0")).join("");
  }

  // ── Osoby ───────────────────────────────────────────────────────────────
  // Rozbiór „Adam Kamiński”, „Kamiński Adam”, „A. Kamiński”, „KAMIŃSKI”, „Anna Kowalska-Nowak”.
  function parsePerson(text) {
    const s = String(text ?? "").trim();
    if (!s) return null;
    const tokens = s.split(/(\s+)/);
    const words = [];
    tokens.forEach((tk, i) => { if (!/^\s+$/.test(tk) && tk) words.push({ w: tk, i }); });
    if (!words.length || words.length > 4) return null;
    const res = { tokens, first: null, firstIdx: -1, initialIdx: -1, surInitialIdx: -1, surnameIdx: [], gender: null };
    words.forEach(({ w, i }) => {
      const bare = w.replace(/[.,]$/, "");
      if (/^\p{Lu}\.?$/u.test(w) && words.length > 1) { if (res.initialIdx < 0 && res.firstIdx < 0) res.initialIdx = i; else if (res.surInitialIdx < 0) res.surInitialIdx = i; return; }
      if (res.firstIdx < 0 && isFirst(bare) && !(res.initialIdx >= 0 && knownSurname(bare))) {
        // „G. Faustyna” — inicjał przed imieniem to inicjał NAZWISKA
        if (res.initialIdx >= 0 && !res.surnameIdx.length) { res.surInitialIdx = res.initialIdx; res.initialIdx = -1; }
        res.firstIdx = i; res.first = bare; res.gender = firstGender(bare); return;
      }
      if (/^[\p{L}'-]+[.,]?$/u.test(w)) res.surnameIdx.push(i);
    });
    // „Jan Maria Kowalski” — drugie imię też jest imieniem; zostawiamy jako nazwisko-część,
    // podmiana i tak idzie słowo po słowie.
    if (!res.surnameIdx.length && res.firstIdx < 0) return null;
    if (!res.gender) {
      const last = res.surnameIdx.length ? tokens[res.surnameIdx[res.surnameIdx.length - 1]] : "";
      res.gender = /(ska|cka|dzka)$/i.test(last) ? "f" : /(ski|cki|dzki)$/i.test(last) ? "m" : null;
    }
    return res;
  }
  function caseLike(src, out) {
    if (src === src.toLocaleUpperCase("pl") && src !== src.toLocaleLowerCase("pl") && src.length > 1) return out.toLocaleUpperCase("pl");
    if (src === src.toLocaleLowerCase("pl")) return out.toLocaleLowerCase("pl");
    return out;
  }
  function inflectSurname(base, gender) {
    if (gender === "f") return base.replace(/ski$/, "ska").replace(/cki$/, "cka").replace(/dzki$/, "dzka");
    return base;
  }
  const asciiFold = (s) => fold(s).replace(/[^a-z0-9]+/g, "");

  // ── Sesja ───────────────────────────────────────────────────────────────
  // opts: { seed, methods:{type→method}, rules, dateShiftDays, labels:{type→"Osoba"}, generalize }
  function createSession(opts = {}) {
    const seed = opts.seed ?? ((Math.random() * 2 ** 32) >>> 0);
    const salt = opts.salt ?? `${seed}:${Math.random()}`;
    const methods = { ...DEFAULT_METHODS, ...(opts.methods || {}) };
    const compiled = compileRules(opts.rules);
    const labels = opts.labels || {};
    const dateShiftDays = opts.dateShiftDays ?? (() => { const r = rng("shift", seed)(); const d = 30 + Math.floor(r * 335); return r < 0.5 ? -d : d; })();
    const pseudoCounters = {};
    const pseudoMap = new Map(); // type|key → numer
    const memo = new Map(); // type|method|key → wynik
    const leakTerms = new Set();
    const fakeFirstMap = new Map();
    const fakeSurMap = new Map();
    const usedFakeFull = new Set();
    // pełne imię znane dla „inicjał + nazwisko” (A. Kamiński → Adam Kamiński), gdy jednoznaczne
    const fullByInitial = new Map();

    const learned = { first: new Map(), sur: new Set(), realFirst: new Set() };
    const activate = () => { LEARNED = learned; };
    activate();
    const surByFirstIni = new Map(); // „faustyna|g” → „grabowska”
    function learnPerson(text) {
      activate();
      let p = parsePerson(text);
      if (!p) return;
      // „Faustyna Grabowska”: żadne słowo nie jest w słowniku imion → imię = słowo bez końcówki
      // nazwiska (zwykle pierwsze: kolejność „Imię Nazwisko”)
      if (p.firstIdx < 0 && p.initialIdx < 0 && p.surnameIdx.length === 2) {
        const [a, b] = p.surnameIdx.map((i) => p.tokens[i].replace(/[.,]$/, ""));
        const strongA = SURNAME_SUFFIX_STRONG.test(a) || knownSurname(a);
        const strongB = SURNAME_SUFFIX_STRONG.test(b) || knownSurname(b);
        const first = strongA && !strongB ? b : a;
        if (!knownSurname(first) || (strongA && strongB)) {
          learned.first.set(norm(first), /a$/i.test(first) && !/^(kuba|barnaba)$/i.test(first) ? "f" : "m");
          p = parsePerson(text);
        }
      }
      if (!p) return;
      p.surnameIdx.forEach((i) => { const w = p.tokens[i].replace(/[.,]$/, ""); if (w.length >= 3 && /^\p{Lu}/u.test(w)) learned.sur.add(norm(w)); });
      if (p.first) learned.realFirst.add(norm(p.first));
      if (p.firstIdx < 0 || !p.surnameIdx.length) return;
      const surJoined = p.surnameIdx.map((i) => norm(p.tokens[i].replace(/[.,]$/, ""))).join(" ");
      const k2 = `${norm(p.first)}|${surJoined[0]}`;
      const prev2 = surByFirstIni.get(k2);
      if (prev2 === undefined) surByFirstIni.set(k2, surJoined);
      else if (prev2 !== surJoined) surByFirstIni.set(k2, null);
      const sur = p.surnameIdx.map((i) => norm(p.tokens[i].replace(/[.,]$/, ""))).join(" ");
      const k = `${sur}|${norm(p.first)[0]}`;
      const prev = fullByInitial.get(k);
      if (prev === undefined) fullByInitial.set(k, norm(p.first));
      else if (prev !== norm(p.first)) fullByInitial.set(k, null); // dwie osoby — niejednoznaczne
    }

    function personKey(text) {
      const p = parsePerson(text);
      if (!p) return `raw:${norm(text)}`;
      let sur = p.surnameIdx.map((i) => norm(p.tokens[i].replace(/[.,]$/, ""))).sort().join(" ");
      if (!sur && p.surInitialIdx >= 0 && p.firstIdx >= 0) {
        const ini = norm(p.tokens[p.surInitialIdx])[0];
        sur = surByFirstIni.get(`${norm(p.first)}|${ini}`) || `${ini}.`;
      }
      if (p.firstIdx >= 0) return `${sur}|${norm(p.first)}`;
      if (p.initialIdx >= 0) {
        const ini = norm(p.tokens[p.initialIdx])[0];
        const full = fullByInitial.get(`${p.surnameIdx.map((i) => norm(p.tokens[i].replace(/[.,]$/, ""))).join(" ")}|${ini}`);
        return full ? `${sur}|${full}` : `${sur}|${ini}.`;
      }
      return `${sur}|`;
    }

    function pseudoNumber(type, key) {
      const k = `${type}|${key}`;
      if (!pseudoMap.has(k)) {
        pseudoCounters[type] = (pseudoCounters[type] || 0) + 1;
        pseudoMap.set(k, pseudoCounters[type]);
      }
      return pseudoMap.get(k);
    }
    const pad = (n) => String(n).padStart(3, "0");
    const label = (type) => labels[type] || type.toUpperCase();

    function fakeFirst(realFirst, gender) {
      const k = norm(realFirst);
      if (!fakeFirstMap.has(k)) {
        const list = (gender || FIRST.get(k) || (/a$/i.test(realFirst) ? "f" : "m")) === "f" ? FEM_FIRST : MALE_FIRST;
        const r = rng(`first:${k}`, seed);
        // jak przy nazwiskach: nie bierzemy imienia, które nosi ktoś w tym pliku („S. Dawid” z
        // fałszywego Dawida wyglądałby jak prawdziwy Dawid S. — test szczelności to łapał)
        let i = Math.floor(r() * list.length);
        for (let tries = 0; tries < list.length && (norm(list[i]) === k || learned.realFirst.has(norm(list[i]))); tries++) i = (i + 1) % list.length;
        fakeFirstMap.set(k, list[i]);
      }
      return fakeFirstMap.get(k);
    }
    function fakeSurnameBase(real) {
      const k = norm(real).replace(/(ska|cka|dzka)$/, (m) => m.replace(/a$/, "i"));
      if (!fakeSurMap.has(k)) {
        const r = rng(`sur:${k}`, seed);
        // nie bierzemy nazwiska, które naprawdę jest w pliku (fałszywy „Nowak” obok prawdziwej Anny Nowak myli)
        const real = (x) => { const n = norm(x); return n === k || learned.sur.has(n) || learned.sur.has(n.replace(/i$/, "a")); };
        let i = Math.floor(r() * FAKE_SURNAMES.length);
        for (let tries = 0; tries < FAKE_SURNAMES.length && real(FAKE_SURNAMES[i]); tries++) i = (i + 1) % FAKE_SURNAMES.length;
        fakeSurMap.set(k, FAKE_SURNAMES[i]);
      }
      return fakeSurMap.get(k);
    }

    // Zamiana osoby słowo po słowie, z zachowaniem formy (kolejność, inicjał, wielkość liter).
    function personFake(text) {
      const p = parsePerson(text);
      if (!p) return `${fakeFirst("x" + norm(text), null)} ${fakeSurnameBase(norm(text))}`;
      const key = personKey(text);
      const fullFirst = key.split("|")[1];
      const gender = p.gender || (fullFirst && firstGender(fullFirst)) || null;
      const tokens = p.tokens.slice();
      if (p.firstIdx >= 0) tokens[p.firstIdx] = caseLike(p.first, fakeFirst(p.first, gender)) + tokens[p.firstIdx].slice(p.first.length);
      if (p.initialIdx >= 0) {
        const src = tokens[p.initialIdx];
        const base = fullFirst && !fullFirst.endsWith(".") ? fakeFirst(fullFirst, gender) : fakeFirst(`ini:${src}`, gender);
        tokens[p.initialIdx] = caseLike(src, `${base[0]}${src.endsWith(".") ? "." : ""}`);
      }
      if (p.surInitialIdx >= 0) {
        const src = tokens[p.surInitialIdx];
        const surFull = key.split("|")[0];
        const base = surFull && !surFull.endsWith(".") ? fakeSurnameBase(surFull) : fakeSurnameBase(`ini:${src}`);
        tokens[p.surInitialIdx] = caseLike(src, `${base[0]}${src.endsWith(".") ? "." : ""}`);
      }
      p.surnameIdx.forEach((i) => {
        const raw = tokens[i];
        const punct = (raw.match(/[.,]$/) || [""])[0];
        const word = punct ? raw.slice(0, -1) : raw;
        const out = word.split("-").map((part) => caseLike(part, inflectSurname(fakeSurnameBase(part), /(ska|cka|dzka)$/i.test(part) ? "f" : /(ski|cki|dzki)$/i.test(part) ? "m" : gender))).join("-");
        tokens[i] = out + punct;
      });
      return tokens.join("");
    }
    function personMask(text) {
      const p = parsePerson(text);
      const words = String(text).trim().split(/\s+/);
      if (!p) return words.map((w) => `${w[0]}.`).join(" ");
      return words.map((w) => (/^\p{L}\.$/u.test(w) ? w : `${w[0]}.`)).join(" ");
    }

    // Fałszywe numery z poprawną sumą kontrolną, w formacie oryginału.
    function reformat(src, newDigits) {
      let i = 0;
      return src.replace(/\d/g, () => newDigits[i++] ?? "0");
    }
    function randDigits(r, n, firstNonZero) {
      let s = "";
      for (let i = 0; i < n; i++) s += String(Math.floor(r() * 10));
      if (firstNonZero && s[0] === "0") s = String(1 + Math.floor(r() * 9)) + s.slice(1);
      return s;
    }
    function fakePesel(src, r) {
      const p = digits(src);
      const d = peselDate(p) || { y: 1980, m: 1, d: 1 };
      const m = 1 + Math.floor(r() * 12);
      const day = 1 + Math.floor(r() * 28);
      const mCode = m + (d.y >= 2000 && d.y < 2100 ? 20 : d.y >= 1800 && d.y < 1900 ? 80 : 0);
      const sex = (+p[9]) % 2;
      let body = String(d.y % 100).padStart(2, "0") + String(mCode).padStart(2, "0") + String(day).padStart(2, "0") + randDigits(r, 3);
      body += String(Math.floor(r() * 5) * 2 + sex);
      return reformat(src, body + peselCheck(body));
    }
    function fakeNip(src, r) {
      for (;;) {
        const b = randDigits(r, 9, true);
        const c = nipCheck(b);
        if (c !== 10 && !/^0{3}/.test(b)) return reformat(src, b + c);
      }
    }
    function fakeRegon(src, r) {
      const n = digits(src).length;
      const b = randDigits(r, n - 1, true);
      return reformat(src, b + regonCheck(b));
    }
    function fakeIban(src, r) {
      const up = String(src).toUpperCase();
      const bban = randDigits(r, 24);
      const check = String(98 - mod97(ibanNumeric(`PL00${bban}`))).padStart(2, "0");
      const all = check + bban;
      if (/^\d/.test(up.replace(/\s/g, ""))) return reformat(src, all);
      let i = 0;
      return src.replace(/^PL/i, (m) => m).replace(/\d/g, () => all[i++] ?? "0");
    }
    function fakeCard(src, r) {
      const n = digits(src).length;
      let b = "4" + randDigits(r, n - 2);
      for (let c = 0; c < 10; c++) if (luhnValid(b + c)) return reformat(src, b + c);
      return reformat(src, b + "0");
    }
    function fakeIdCard(src, r) {
      const L = "ABCDEFGHIJKLMNOPRSTUWYZ";
      const letters = Array.from({ length: 3 }, () => L[Math.floor(r() * L.length)]).join("");
      const tail = randDigits(r, 5);
      const tmp = `${letters}0${tail}`;
      const v = `${letters}${idCardCheck(tmp)}${tail}`;
      return /\s/.test(src) ? `${v.slice(0, 3)} ${v.slice(3)}` : v;
    }
    function fakePassport(src, r) {
      const L = "ABCDEFGHIJKLMNOPRSTUWYZ";
      const letters = L[Math.floor(r() * L.length)] + L[Math.floor(r() * L.length)];
      const tail = randDigits(r, 6);
      const v = `${letters}${passportCheck(`${letters}0${tail}`)}${tail}`;
      return /\s/.test(src) ? `${v.slice(0, 2)} ${v.slice(2)}` : v;
    }
    function fakeKw(src, r) {
      const m = String(src).toUpperCase().match(/^([A-Z]{2}\d[A-Z])\/(\d{8})\/(\d)$/);
      const court = m ? m[1] : "WA1M";
      const num = randDigits(r, 8);
      return `${court}/${num}/${kwCheck(court, num)}`;
    }
    function fakePhone(src, r) {
      const d = digits(src);
      const keepCountry = d.length > 9 ? d.slice(0, d.length - 9) : "";
      const body = String(5 + Math.floor(r() * 4)) + randDigits(r, 8);
      return reformat(src, keepCountry + body);
    }
    function fakeEmail(src, r) {
      const f = FEM_FIRST.concat(MALE_FIRST);
      const a = f[Math.floor(r() * f.length)];
      const b = FAKE_SURNAMES[Math.floor(r() * FAKE_SURNAMES.length)];
      return `${asciiFold(a)}.${asciiFold(b)}${Math.floor(r() * 90 + 10)}@example.com`;
    }
    // Ta sama budowa co oryginał: prefiks (albo jego brak), numer mieszkania, kod i miejscowość.
    function fakeAddress(src, r) {
      const str = String(src).trim();
      const pm = str.match(RE_ADDR_PREFIX_START);
      const prefix = pm ? (/\s$/.test(pm[0]) ? pm[0] : `${pm[0]} `) : /\d/.test(str) ? "" : "ul. ";
      const own = norm(str);
      let k = Math.floor(r() * STREETS.length);
      while (own.includes(norm(STREETS[k]))) k = (k + 1) % STREETS.length; // nie ta sama ulica co prawdziwa
      const street = STREETS[k];
      const no = 1 + Math.floor(r() * 120);
      const flatM = str.match(/\d\s?\/\s?\d|(?:m\.?|lok\.?|lokal)\s*\d/);
      const flat = !flatM ? "" : /\//.test(flatM[0]) ? `/${1 + Math.floor(r() * 60)}` : ` m. ${1 + Math.floor(r() * 60)}`;
      const pc = str.match(/(,?\s*)\d{2}-\d{3}(\s+\p{Lu})?/u);
      const tail = pc ? `${pc[1] || " "}${String(10 + Math.floor(r() * 89))}-${randDigits(r, 3)}${pc[2] ? ` ${CITIES.filter((c) => !own.includes(norm(c)))[Math.floor(r() * (CITIES.length - 1))] || CITIES[0]}` : ""}` : "";
      return /\d/.test(str) ? `${prefix}${street} ${no}${flat}${tail}` : `${prefix}${street}`;
    }
    function fakeAlnum(src, r) {
      return String(src).replace(/[A-Z]/g, () => "ABCDEFGHJKLMNPRSTUVWXYZ"[Math.floor(r() * 23)]).replace(/[a-z]/g, () => "abcdefghjkmnprstuvwxyz"[Math.floor(r() * 22)]).replace(/\d/g, () => String(Math.floor(r() * 10)));
    }
    function fakePlate(src, r) {
      const m = String(src).match(/^([A-Z]{2,3})(\s?)(.*)$/);
      return m ? `${m[1].slice(0, 2)}${m[1].length === 3 ? "ABCDEFGHJKLMNPRSTUVWXYZ"[Math.floor(r() * 23)] : ""}${m[2]}${fakeAlnum(m[3], r)}` : fakeAlnum(src, r);
    }

    function maskValue(type, s) {
      const str = String(s);
      if (type === "person") return personMask(str);
      if (type === "email") {
        const at = str.indexOf("@");
        return at > 0 ? `${str[0]}***${str.slice(at)}` : "***";
      }
      if (type === "address") return str.replace(/(ul\.|ulica|al\.|aleja|aleje|os\.|osiedle|pl\.|plac)?(\s*)(.*)/i, (m, p, sp, rest) => `${p || ""}${sp}${rest.replace(/[\p{L}\d]/gu, "*")}`);
      if (type === "postcode") return str.replace(/\d{3}$/, "***");
      if (type === "ip") return str.replace(/\.\d+\.\d+$/, ".*.*");
      // numery: zostają ostatnie 3 (karta/konto — 4), reszta gwiazdki, separatory zostają
      const keep = type === "iban" || type === "card" ? 4 : type === "pesel" || type === "nip" || type === "regon" || type === "phone" ? 3 : 2;
      let seen = 0;
      const total = (str.match(/[\p{L}\d]/gu) || []).length;
      return str.replace(/[\p{L}\d]/gu, (ch) => (++seen > total - keep ? ch : "*"));
    }

    function generalizeValue(type, s) {
      const str = String(s);
      if (type === "postcode") return str.replace(/-\d{3}$/, "-xxx");
      if (type === "pesel") { const d = peselDate(digits(str)); return d ? String(d.y) : "***"; }
      // adres → miejscowość (gdy jest po kodzie), inaczej sama ulica bez numeru („ul. Długa”)
      if (type === "address") {
        const town = str.match(/\d{2}-\d{3}\s+(\p{Lu}\p{Ll}+(?:[ -]\p{Lu}\p{Ll}+)?)\s*$/u);
        if (town) return town[1];
        const m = str.match(ADDR_RE_CELL);
        if (m && /\d/.test(str)) return `${(str.match(RE_ADDR_PREFIX_START) || [""])[0]}${m[1]}`.trim();
        return `[${label("address")}]`;
      }
      if (type === "phone") return reformat(str, digits(str).slice(0, digits(str).length - 6) + "000000");
      if (type === "ip") return str.replace(/\.\d+$/, ".0");
      if (type === "person") return label("person");
      if (type === "email") { const at = str.indexOf("@"); return at > 0 ? `***${str.slice(at)}` : "***"; }
      return `[${label(type)}]`;
    }

    // Zamiana jednego trafienia (fragmentu tekstu). Zwraca nowy tekst.
    function replaceText(type, value, method) {
      const m = method || methods[type] || "redact";
      if (m === "keep") return value;
      const key = type === "person" ? personKey(value) : type === "email" || type === "custom" ? norm(value) : norm(value).replace(/[\s-]/g, "");
      // osoba fałszywa/maska zależy od formy zapisu (inicjał, WIELKIE) — spójność dają mapy imion i nazwisk
      const mk = type === "person" && (m === "fake" || m === "mask") ? `${type}|${m}|=${String(value)}` : `${type}|${m}|${key}`;
      if (memo.has(mk)) { addLeak(type, value); return memo.get(mk); }
      let out;
      const r = rng(`${type}:${key}`, seed);
      switch (m) {
        case "remove": out = ""; break;
        case "redact": out = `[${label(type)}]`; break;
        case "mask": out = maskValue(type, value); break;
        case "hash": out = sha256Hex(`${salt}|${type}|${key}`).slice(0, 12); break;
        case "pseudo": out = type === "email" ? `osoba${pad(pseudoNumber(type, key))}@example.com` : `${label(type)} ${pad(pseudoNumber(type, key))}`; break;
        case "generalize": out = generalizeValue(type, value); break;
        case "fake": out = fakeFor(type, value, r); break;
        default: out = `[${label(type)}]`;
      }
      memo.set(mk, out);
      addLeak(type, value);
      return out;
    }
    function fakeFor(type, value, r) {
      switch (type) {
        case "person": {
          let out = personFake(value);
          // dwie różne osoby nie mogą dostać tej samej fałszywej (rzadkie zderzenie losowania)
          const k = personKey(value);
          for (let i = 0; usedFakeFull.has(`${norm(out)}`) && !usedFakeFull.has(`${norm(out)}#${k}`) && i < 5; i++) out = `${out} ${"ABCDEFGH"[i]}.`;
          usedFakeFull.add(norm(out)); usedFakeFull.add(`${norm(out)}#${k}`);
          return out;
        }
        case "pesel": return fakePesel(value, r);
        case "nip": return fakeNip(value, r);
        case "regon": return fakeRegon(value, r);
        case "iban": return fakeIban(value, r);
        case "card": return fakeCard(value, r);
        case "idcard": return fakeIdCard(value.toUpperCase(), r);
        case "passport": return fakePassport(value.toUpperCase(), r);
        case "kw": return fakeKw(value, r);
        case "phone": return fakePhone(value, r);
        case "email": return fakeEmail(value, r);
        case "address": return fakeAddress(value, r);
        case "postcode": return `${String(value).slice(0, 2)}-${randDigits(r, 3)}`;
        case "place": return CITIES[Math.floor(r() * CITIES.length)];
        case "plate": return fakePlate(value, r);
        case "ip": return `192.0.2.${1 + Math.floor(r() * 254)}`;
        default: return fakeAlnum(value, r);
      }
    }
    function addLeak(type, value) {
      const v = String(value).trim();
      if (v.length >= 3) leakTerms.add(v);
      if (type === "person") {
        const p = parsePerson(v);
        if (p) p.surnameIdx.forEach((i) => { const w = p.tokens[i].replace(/[.,]$/, ""); if (w.length >= 4) leakTerms.add(w); });
      }
      if (/^[\d\s+()-]+$/.test(v) && digits(v).length >= 7) leakTerms.add(digits(v));
    }

    // Wartość CAŁEJ komórki w kolumnie typu `type` (np. kolumna „Właściciel”).
    // value: string | number | Date. Zwraca { value, type:"string"|"number"|"date"|null(=wyczyść) }.
    function replaceWholeCell(type, value, method) {
      const m = method || methods[type] || "redact";
      if (m === "keep") return null;
      if (value instanceof Date) {
        if (m === "shift") return { value: new Date(value.getTime() + dateShiftDays * 86400000), type: "date" };
        if (m === "generalize" || m === "mask") return { value: type === "dob" ? value.getFullYear() : `${value.getFullYear()}-${String(value.getMonth() + 1).padStart(2, "0")}`, type: type === "dob" ? "number" : "string" };
        if (m === "remove") return { value: null, type: null };
        if (m === "fake") { const r = rng(`date:${value.getTime()}`, seed); return { value: new Date(value.getTime() + Math.round((r() - 0.5) * 2 * 365) * 86400000), type: "date" }; }
        const out = replaceText(type, value.toISOString().slice(0, 10), m);
        return { value: out, type: "string" };
      }
      if (typeof value === "number") {
        const s = Number.isInteger(value) ? String(value) : String(value);
        if (m === "remove") return { value: null, type: null };
        if (m === "generalize" && !["pesel", "nip", "regon", "phone", "iban", "card"].includes(type)) {
          const mag = Math.pow(10, Math.max(0, Math.floor(Math.log10(Math.abs(value) || 1)) - 1));
          return { value: Math.round(value / mag) * mag, type: "number" };
        }
        const out = replaceText(type, s, m);
        if (m === "fake" && /^\d+$/.test(out) && out.length < 16 && out[0] !== "0") return { value: Number(out), type: "number" };
        return { value: out, type: "string" };
      }
      const str = String(value);
      if (!str.trim()) return null;
      if (m === "remove") return { value: null, type: null };
      if (m === "shift" && type === "dob") return null;
      return { value: replaceText(type, str.trim(), m), type: "string" };
    }

    // Tekst z trafieniami w środku: zamienia tylko trafienia (metoda wg typu).
    function replaceInText(text, hits, methodFor) {
      let out = "";
      let pos = 0;
      for (const h of hits) {
        out += text.slice(pos, h.start);
        out += replaceText(h.type, h.value, methodFor ? methodFor(h.type) : undefined);
        pos = h.end;
      }
      return out + text.slice(pos);
    }

    // Przepisanie dowolnego tekstu (wykresy, tabele przestawne, nagłówki stron): najpierw
    // znane oryginały w całości, potem wykrywanie.
    function rewriteString(s) {
      activate();
      const str = String(s ?? "");
      if (!str) return str;
      const exact = exactMap.get(norm(str));
      if (exact !== undefined) return exact;
      // trafienia liczone na ORYGINALE (nie na wyniku — fałszywa osoba nie może być zamieniona drugi raz)
      const spans = [];
      if (knownRe) { knownRe.lastIndex = 0; for (const m of str.matchAll(knownRe)) spans.push({ start: m.index, end: m.index + m[0].length, out: knownMap.get(norm(m[0])) ?? m[0] }); }
      for (const h of findInText(str, compiled)) {
        if (spans.some((x) => h.start < x.end && h.end > x.start)) continue;
        spans.push({ start: h.start, end: h.end, out: replaceText(h.type, h.value) });
      }
      spans.sort((a, b) => a.start - b.start);
      let out = "", pos = 0;
      const gap = (t) => (surnameRe ? t.replace(surnameRe, (m) => replaceText("person", m)) : t);
      for (const x of spans) { out += gap(str.slice(pos, x.start)) + x.out; pos = x.end; }
      return out + gap(str.slice(pos));
    }
    let surnameRe = null; // poznane nazwiska z pliku — ostatnia linia obrony w śladach poza komórkami
    const exactMap = new Map();
    const knownMap = new Map();
    let knownRe = null;
    function remember(original, replaced) {
      const k = norm(original);
      if (!k || k === norm(replaced)) return;
      exactMap.set(k, replaced);
      if (k.length >= 4) { knownMap.set(k, replaced); knownRe = null; }
    }
    function freezeKnown() {
      const keys = [...knownMap.keys()].sort((a, b) => b.length - a.length);
      knownRe = keys.length ? new RegExp(`(?<![\\p{L}\\p{N}])(?:${keys.map(escapeRe).join("|")})(?![\\p{L}\\p{N}])`, "giu") : null;
      const surs = [...learned.sur].filter((x) => x.length >= 4 && !FIRST.has(x)).sort((a, b) => b.length - a.length);
      surnameRe = surs.length ? new RegExp(`(?<![\\p{L}\\p{N}])(?:${surs.map(escapeRe).join("|")})(?![\\p{L}\\p{N}])`, "giu") : null;
    }

    return {
      seed, salt, methods, compiled, dateShiftDays, leakTerms,
      learnPerson, personKey, rewriteString, remember, freezeKnown, activate, learned,
      replaceText: (...a) => { activate(); return replaceText(...a); },
      replaceWholeCell: (...a) => { activate(); return replaceWholeCell(...a); },
      replaceInText: (...a) => { activate(); return replaceInText(...a); },
      findInText: (text, ctx) => { activate(); return findInText(text, compiled, ctx); },
    };
  }

  // Szukanie wycieków: które z oryginałów zostały w tekście (bez wielkości liter).
  function findLeaks(text, terms) {
    const low = String(text).toLocaleLowerCase("pl");
    const found = [];
    for (const t of terms) if (low.includes(String(t).toLocaleLowerCase("pl"))) found.push(t);
    return found;
  }

  return {
    TYPES, METHODS, DEFAULT_METHODS, norm, fold,
    validators: { peselValid, nipValid, regonValid, ibanValid, luhnValid, idCardValid, passportValid, kwValid },
    resetLearned: () => { LEARNED = { first: new Map(), sur: new Set() }; },
    findInText, compileRules, headerHint, profileColumn, personCellScore, addressCellScore, parsePerson, createSession, findLeaks, sha256Hex,
  };
})();

if (typeof module !== "undefined" && module.exports) module.exports = SwbAnon;
