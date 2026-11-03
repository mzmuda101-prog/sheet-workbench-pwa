#!/usr/bin/env node
// Testy silnika anonimizacji (app/anonymize-core.js) — bez przeglądarki.
const A = require("../app/anonymize-core.js");

let fails = 0, passes = 0;
function ok(cond, msg) { if (cond) passes++; else { fails++; console.log("  ✗", msg); } }
function eq(a, b, msg) { ok(a === b, `${msg} — jest ${JSON.stringify(a)}, oczekiwano ${JSON.stringify(b)}`); }
const types = (s) => A.findInText(s).map((h) => `${h.type}:${h.value}`);

// ── walidatory ─────────────────────────────────────────────
const V = A.validators;
ok(V.peselValid("44051401359"), "PESEL poprawny");
ok(!V.peselValid("44051401358"), "PESEL zła suma");
ok(V.peselValid("02270803624"), "PESEL 2000+");
ok(!V.peselValid("44133101359"), "PESEL zły miesiąc");
ok(V.nipValid("526-000-12-46"), "NIP z kreskami");
ok(V.nipValid("5260001246"), "NIP goły");
ok(!V.nipValid("5260001247"), "NIP zła suma");
ok(V.regonValid("123456785"), "REGON 9");
ok(V.regonValid("12345678512347"), "REGON 14");
ok(V.ibanValid("PL61 1090 1014 0000 0712 1981 2874"), "IBAN PL");
ok(V.ibanValid("61109010140000071219812874"), "NRB bez PL");
ok(!V.ibanValid("61109010140000071219812875"), "NRB zła suma");
ok(V.luhnValid("4111 1111 1111 1111"), "karta Luhn");
ok(V.idCardValid("ABA300000"), "dowód ABA300000");
ok(!V.idCardValid("ABA300001"), "dowód zła cyfra");

// ── wykrywanie w tekście ────────────────────────────────────
const t1 = types("Kontakt: jan.kowalski@firma.pl, tel. 600 123 456, PESEL 44051401359");
ok(t1.includes("email:jan.kowalski@firma.pl"), "e-mail w tekście");
ok(t1.includes("phone:600 123 456"), "telefon w tekście");
ok(t1.includes("pesel:44051401359"), "PESEL w tekście");
ok(types("Rachunek PL61 1090 1014 0000 0712 1981 2874").some((x) => x.startsWith("iban:")), "IBAN w tekście");
ok(types("ul. Lipowa 12/3, 62-200 Gniezno").some((x) => x.startsWith("address:ul. Lipowa 12/3")), "adres");
// kod i miejscowość po adresie należą do adresu (zamiennik ma tę samą budowę); sam kod osobno
ok(types("ul. Lipowa 12/3, 62-200 Gniezno").includes("address:ul. Lipowa 12/3, 62-200 Gniezno"), "adres z kodem i miejscowością");
ok(types("kod pocztowy 62-200").includes("postcode:62-200"), "kod pocztowy");
ok(types("Spotkanie z Adamem? Rozmawiał Jan Kowalski wczoraj").includes("person:Jan Kowalski"), "osoba w zdaniu");
ok(types("Odebrał B. Bielak").length === 0 || types("Odebrał B. Bielak").includes("person:B. Bielak"), "inicjał + nazwisko (słownik)");
ok(types("Pani Maślankiewicz dzwoniła").some((x) => x.startsWith("person:")), "Pani + nazwisko");
eq(types("Suma 1234 zł, faktura 2024/05").length, 0, "zwykły tekst bez trafień");
eq(types("Kwota 12345678901").length, 0, "11 cyfr bez sumy PESEL — nie PESEL");
ok(types("KW PO1P/00012345/6").some((x) => x.startsWith("kw:")), "księga wieczysta");
ok(types("Auto WA 12345").some((x) => x.startsWith("plate:")), "rejestracja");

// wyjątki i własne słowa
const c = A.compileRules({ exceptions: ["Jan Kowalski"], terms: ["Projekt Sowa"] });
eq(A.findInText("Jan Kowalski i Projekt Sowa", c).map((h) => h.type).join(","), "custom", "wyjątek + własne słowo");

// ── nagłówki i kolumny ─────────────────────────────────────
eq(A.headerHint("Imię i Nazwisko"), "person", "nagłówek osoby");
eq(A.headerHint("Nr telefonu"), "phone", "nagłówek telefonu");
eq(A.headerHint("PESEL"), "pesel", "nagłówek PESEL");
eq(A.headerHint("Kwota"), null, "nagłówek zwykły");
const prof = A.profileColumn("Pracownik", ["P. Pilipczuk", "S. Pilipczuk", "Adam Nowak"].map((v) => ({ v })));
eq(prof.type, "person", "kolumna osób z nagłówka");
const prof2 = A.profileColumn("Kolumna X", ["Adam Kamiński", "Anna Nowak", "Piotr Zieliński", "Ewa Bąk"].map((v) => ({ v })));
eq(prof2.type, "person", "kolumna osób z treści");
const prof3 = A.profileColumn("Status", ["W trakcie", "Zamknięte", "W trakcie", "Spóźnione"].map((v) => ({ v })));
eq(prof3.type, null, "kolumna statusów — nie osoby");

// ── sesja: spójność i formy osoby ──────────────────────────
const s = A.createSession({ seed: 7, labels: { person: "Osoba" } });
["Adam Kamiński", "Anna Kamińska", "Piotr Nowak"].forEach(s.learnPerson);
const p1 = s.replaceText("person", "Adam Kamiński", "pseudo");
const p2 = s.replaceText("person", "A. Kamiński", "pseudo");
const p3 = s.replaceText("person", "Kamiński Adam", "pseudo");
eq(p1, "Osoba 001", "pseudonim 1");
eq(p2, p1, "inicjał = ta sama osoba");
eq(p3, p1, "odwrócona kolejność = ta sama osoba");
ok(s.replaceText("person", "Anna Kamińska", "pseudo") !== p1, "inna osoba = inny pseudonim");
const f1 = s.replaceText("person", "Adam Kamiński", "fake");
const f2 = s.replaceText("person", "A. Kamiński", "fake");
const f3 = s.replaceText("person", "ADAM KAMIŃSKI", "fake");
ok(f1 !== "Adam Kamiński" && /^\p{Lu}\p{Ll}+ \p{Lu}\p{Ll}+$/u.test(f1), `fałszywa osoba: ${f1}`);
eq(f2, `${f1[0]}. ${f1.split(" ")[1]}`, "fałszywa w formie inicjału");
eq(f3, f1.toLocaleUpperCase("pl"), "fałszywa WIELKIMI");
const fem = s.replaceText("person", "Anna Kamińska", "fake");
ok(!/ski$/.test(fem), `kobieta nie dostaje końcówki -ski: ${fem}`);
// numery fałszywe są poprawne
const fp = s.replaceText("pesel", "44051401359", "fake");
ok(V.peselValid(fp) && fp !== "44051401359", `fałszywy PESEL poprawny: ${fp}`);
ok(V.nipValid(s.replaceText("nip", "526-000-12-46", "fake")), "fałszywy NIP poprawny");
ok(/^\d{3}-\d{3}-\d{2}-\d{2}$/.test(s.replaceText("nip", "526-000-12-46", "fake")), "fałszywy NIP w formacie");
ok(V.ibanValid(s.replaceText("iban", "PL61 1090 1014 0000 0712 1981 2874", "fake")), "fałszywy IBAN poprawny");
ok(V.idCardValid(s.replaceText("idcard", "ABA300000", "fake")), "fałszywy dowód poprawny");
eq(s.replaceText("pesel", "44051401359", "mask"), "********359", "maska PESEL");
eq(s.replaceText("email", "jan.kowalski@firma.pl", "mask"), "j***@firma.pl", "maska e-mail");
eq(s.replaceText("postcode", "62-200", "generalize"), "62-xxx", "uogólnienie kodu");
eq(s.replaceText("person", "Adam Kamiński", "mask"), "A. K.", "maska osoby = inicjały");
eq(s.replaceText("pesel", "44051401359", "hash").length, 12, "skrót 12 znaków");
eq(s.replaceText("pesel", "44051401359", "hash"), s.replaceText("pesel", "44 051 401 359".replace(/ /g, ""), "hash"), "skrót spójny");
// cała komórka: data
const d = new Date(1985, 2, 14);
const sh = s.replaceWholeCell("dob", d, "shift");
eq(Math.round((sh.value - d) / 86400000), s.dateShiftDays, "przesunięcie daty o stałą liczbę dni");
eq(s.replaceWholeCell("dob", d, "generalize").value, 1985, "data urodzenia → rok");
eq(s.replaceWholeCell("pesel", 44051401359, "mask").value, "********359", "PESEL jako liczba → maska");
// tekst z trafieniami
const txt = "Dzwonił Jan Kowalski, tel. 600 123 456";
const out = s.replaceInText(txt, s.findInText(txt));
ok(!/Kowalski|600 123 456/.test(out), `tekst bez danych: ${out}`);
// wycieki
ok(s.leakTerms.has("Kamiński"), "nazwisko w terminach wycieku");
eq(A.findLeaks("<t>adam kamiński</t>", ["Adam Kamiński"]).length, 1, "wyciek bez wielkości liter");
// przepisywanie po pamięci
s.remember("Adam Kamiński", "Osoba 001");
s.freezeKnown();
eq(s.rewriteString("Adam Kamiński"), "Osoba 001", "przepisanie dokładne");
eq(s.rewriteString("Raport: Adam Kamiński (C1)"), "Raport: Osoba 001 (C1)", "przepisanie w środku");
eq(A.sha256Hex("abc"), "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad", "SHA-256");

// imiona spoza słownika poznane z kolumny osób + formy pochodne (RODO „Baza danych”)
const s2 = A.createSession({ seed: 3, labels: { person: "Osoba" } });
["Faustyna Grabowska", "Marek Żmuda", "Jolanta Murlak"].forEach(s2.learnPerson);
s2.freezeKnown();
const g1 = s2.replaceText("person", "Faustyna Grabowska", "pseudo");
eq(s2.rewriteString("Grabowska Faustyna"), g1, "odwrócona forma z poznanym imieniem");
eq(s2.rewriteString("G. Faustyna"), g1, "inicjał nazwiska + imię");
ok(!/Żmuda/.test(s2.rewriteString("M. Żmuda")), "inicjał + poznane nazwisko");
ok(!/Murlak/.test(s2.rewriteString("Gr1 J Murlak")), `inicjał bez kropki + poznane nazwisko: ${s2.rewriteString("Gr1 J Murlak")}`);
ok(/^Gr1 /.test(s2.rewriteString("Gr1 J Murlak")), "prefiks grupy zostaje");
const ff = s2.replaceText("person", "Faustyna Grabowska", "fake");
eq(s2.replaceText("person", "G. Faustyna", "fake"), `${ff.split(" ")[1][0]}. ${ff.split(" ")[0]}`, "fałszywa forma „G. Faustyna”");
ok(s2.rewriteString("W trakcie") === "W trakcie", "zwykły tekst nietknięty po nauce");
const s3 = A.createSession({ seed: 3, methods: { person: "fake" } });
["Jan Kowalski"].forEach(s3.learnPerson);
const fk = s3.replaceText("person", "Jan Kowalski", "fake");
s3.remember("Jan Kowalski", fk); s3.freezeKnown();
eq(s3.rewriteString(`Jan Kowalski (C1)`), `${fk} (C1)`, "fałszywa osoba nie jest zamieniana drugi raz");

// etykiety w kolumnie z osobami nie są osobami (zgłoszenie 2026-11-01: „Porządkowi”, „Mikrofony”, „STARSI”)
["Porządkowi", "Nagłośnienie", "Mikrofony", "STARSI", "SŁUDZY", "Data", "Otwieranie"].forEach((w) => eq(A.personCellScore(w), 0, `etykieta „${w}” to nie osoba`));
["Jan Kowalski", "B. Bielak", "Kowalski", "Faustyna Grabowska", "J Murlak"].forEach((w) => ok(A.personCellScore(w) > 0, `„${w}” wygląda na osobę`));

// odmienione, małą literą (notatka w „bazy”, 2026-11-01)
ok(types("pamiętać dodać andrzeja słowikowskiego do audytorium").includes("person:andrzeja słowikowskiego"), "imię+nazwisko w dopełniaczu, małe litery");
ok(types("rozmawiałem z pawłem nowickim wczoraj").includes("person:pawłem nowickim"), "narzędnik, wstawne e");
ok(types("Przekaż Annie Kowalskiej klucze").some((x) => x.startsWith("person:Annie Kowalskiej")), "celownik, wielkie litery");
eq(types("dla polskiego zespołu w trakcie").length, 0, "zwykłe słowa z końcówką -skiego nie są osobą");
eq(types("3 maja polskiego święta").length, 0, "„maja” to nie imię");

// ── adresy (zgłoszenie 2026-11-01: wyrywanie ulic) ─────────────
[
  ["Mieszka przy ul. Długa 5/3 w Krakowie", "address:ul. Długa 5/3"],
  ["al. Jana Pawła II 30", "address:al. Jana Pawła II 30"],
  ["ul. 3 Maja 1", "address:ul. 3 Maja 1"],
  ["ul. gen. Andersa 7a", "address:ul. gen. Andersa 7a"],
  ["os. Słoneczne 4 m. 2", "address:os. Słoneczne 4 m. 2"],
  ["UL. KRÓTKA 4", "address:UL. KRÓTKA 4"],
  ["Adres: ulica Długa", "address:ulica Długa"],
  ["Lipowa 12, 31-100 Kraków", "address:Lipowa 12, 31-100 Kraków"],
  ["ul. Adama Mickiewicza 12 m. 4, 00-950 Warszawa", "address:ul. Adama Mickiewicza 12 m. 4, 00-950 Warszawa"],
].forEach(([s, want]) => ok(types(s).includes(want), `adres: ${s} — jest ${JSON.stringify(types(s))}`));
ok(!types("ul. Adama Mickiewicza 12").some((x) => x.startsWith("person:")), "osoba w nazwie ulicy należy do adresu");
eq(types("Grupa 3 i Rozdział 5, Gr3 Osoba 010").length, 0, "etykiety z numerem to nie adres");
eq(A.addressCellScore("Osoba 010"), 0, "numer z zerem wiodącym to nie numer domu");
eq(A.profileColumn("Teren", ["Długa 5", "Lipowa 12/3", "Krótka 1", "Polna 7a"].map((v) => ({ v }))).type, "address", "kolumna „Długa 5…” = adresy (po treści)");
eq(A.profileColumn("Kolumna", ["Grupa 1", "Grupa 2", "Grupa 3", "Grupa 4"].map((v) => ({ v }))).type, null, "„Grupa 1…4” to nie adresy");
eq(A.profileColumn("Ulica", ["Długa", "Polna"].map((v) => ({ v }))).type, "address", "nagłówek „Ulica”");
{
  const sa = A.createSession({ seed: 7, methods: { address: "fake" } });
  const f1 = sa.replaceText("address", "Lipowa 12, 31-100 Kraków", "fake");
  ok(/^\p{Lu}\p{Ll}+ \d+, \d{2}-\d{3} \p{Lu}/u.test(f1) && !/Lipowa|Kraków/.test(f1), `fałszywy adres ma tę samą budowę — ${f1}`);
  ok(/^ul\. \p{Lu}\p{Ll}+ \d+\/\d+$/u.test(sa.replaceText("address", "ul. Długa 5/3", "fake")), "fałszywy adres z mieszkaniem");
  eq(sa.replaceText("address", "Lipowa 12, 31-100 Kraków", "generalize"), "Kraków", "uogólnienie → miejscowość");
  eq(sa.replaceText("address", "ul. Długa 5/3", "generalize"), "ul. Długa", "uogólnienie → ulica bez numeru");
}

// ── polowanie na błędy 2026-11-03 ─────────────
{
  // remember() po freezeKnown() (autor notatki przy czyszczeniu zipa) nie może gasić znanych oryginałów
  const sk = A.createSession({ seed: 3, methods: { person: "pseudo" } });
  sk.remember("Remontex Hurt", "[FIRMA]");
  sk.freezeKnown();
  sk.remember("jkowal", "Osoba 009");
  eq(sk.rewriteString("Dostawca Remontex Hurt sp."), "Dostawca [FIRMA] sp.", "znany oryginał zamieniany także po kolejnym remember()");
  eq(sk.rewriteString("login jkowal"), "login Osoba 009", "nowy oryginał z remember() też zamieniany");
  // skrót: samotny surogat (uszkodzone emoji) nie wywraca sesji
  const sh = A.createSession({ seed: 3, methods: { person: "hash" } });
  let threw = null;
  try { sh.replaceText("person", "Jan Kowalski \ud83d", "hash"); } catch (e) { threw = e; }
  ok(!threw, `skrót z samotnym surogatem — ${threw && threw.message}`);
}

console.log(`${fails ? "❌" : "✅"} anonymize-core: ${passes} ok, ${fails} błędów`);
process.exit(fails ? 1 : 0);
