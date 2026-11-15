// entry-fixture.js — wspólny plik testowy Wpisywania (openpyxl).
// Arkusz „Prace”, tabela Excela „Prace” A1:E6:
//   Osoba (5 osób, „Ewa Lis” dwa razy), Data (dd.mm.yyyy), Status (lista „wzięta,zdana”, stop),
//   Ilość (liczby), Razem (kolumna obliczeniowa =D2*2).
// Wiersz 3 (2. Ewa Lis) ma puste Data i Status — do uzupełniania.
const path = require("path");
const fs = require("fs");
const { spawnSync } = require("child_process");

const VENV_PY = path.join(__dirname, "..", ".venv/bin/python");
const PY = fs.existsSync(VENV_PY) ? VENV_PY : "python3";

function makeEntryFixture(file) {
  const py = spawnSync(PY, ["-I", "-c", `
import sys, datetime
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo, TableFormula
from openpyxl.worksheet.datavalidation import DataValidation
wb = Workbook(); ws = wb.active; ws.title = "Prace"
ws.append(["Osoba", "Data", "Status", "Ilość", "Razem"])
data = [("Jan Kowalski", datetime.datetime(2026, 3, 1), "wzięta", 2),
        ("Ewa Lis", None, None, 3),
        ("Anna Nowak", datetime.datetime(2026, 3, 3), "zdana", 1),
        ("Ewa Lis", datetime.datetime(2026, 3, 4), "wzięta", 5),
        ("Piotr Zając", datetime.datetime(2026, 3, 5), "zdana", 4)]
for i, (o, d, s, n) in enumerate(data):
    r = i + 2
    ws.cell(r, 1, o); ws.cell(r, 4, n); ws.cell(r, 5, "=D%d*2" % r)
    if s: ws.cell(r, 3, s)
    c = ws.cell(r, 2, d); c.number_format = "dd.mm.yyyy"
ws.cell(3, 2).number_format = "dd.mm.yyyy"
dv = DataValidation(type="list", formula1='"wzięta,zdana"', allow_blank=True, errorStyle="stop")
dv.add("C2:C6"); ws.add_data_validation(dv)
t = Table(displayName="Prace", ref="A1:E6"); t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
t._initialise_columns()
for i, n in enumerate(["Osoba", "Data", "Status", "Ilość", "Razem"]): t.tableColumns[i].name = n
t.tableColumns[4].calculatedColumnFormula = TableFormula(attr_text="D2*2")
ws.add_table(t)
wb.save(sys.argv[1])
`, file], { encoding: "utf8" });
  if (py.status !== 0) throw new Error("openpyxl: " + py.stderr);
}

module.exports = { makeEntryFixture };
