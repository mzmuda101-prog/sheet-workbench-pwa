"""Generuje scripts/fixtures/excel-parity-styles.xlsx (test excel-parity-styles-playwright.js).

Uruchom: .venv/bin/python scripts/fixtures/make-excel-parity-styles.py
Każdy przypadek to coś, co aplikacja kiedyś pokazywała inaczej niż Excel (2026-11-01/02).
"""
import datetime
import os

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Color, Font, GradientFill, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "excel-parity-styles.xlsx")

wb = Workbook()
ws = wb.active
ws.title = "Style"
# wiersz nad nagłówkiem: czerwony z palety (indexed 2) + „Zawijaj tekst”
ws["A1"] = "Plan"
ws["A1"].font = Font(bold=True, size=20)
ws["D1"] = "Aktualizacja 1.1: zakresy bez poprzedniego tygodnia"
ws["D1"].font = Font(size=10, color=Color(indexed=2))
ws["D1"].alignment = Alignment(wrap_text=True)
ws.row_dimensions[1].height = 48.75
ws.column_dimensions["D"].width = 14
for c, h in zip("ABCD", ["Data", "Gradient", "Żółte", "Czerwone"]):
    ws[c + "3"] = h
# gradient MUSI być pierwszym wypełnieniem w pliku — po nim indeksy SheetJS się przesuwały
ws["B4"].fill = GradientFill(type="path", stop=("FFC000", "FFFFFF"))
ws["B4"] = "g"
ws["C4"] = "żółte"
ws["C4"].fill = PatternFill("solid", fgColor="FFFF99")
ws["D4"] = "czerwone"
ws["D4"].fill = PatternFill("solid", fgColor="C00000")
for i in range(6):
    r = 4 + i
    ws.cell(r, 1, datetime.datetime(2026, 9, 8) + datetime.timedelta(days=7 * i)).number_format = "[$-415]d/mmm/yyyy;@"
    if i:
        ws.cell(r, 2, "x%d" % i)
        ws.cell(r, 3, i)
        ws.cell(r, 4, "Osoba %03d" % i)
ws.column_dimensions["A"].width = 13
# scalona notatka: zawijanie, szare tło z palety (indexed 22), ramka górna gruba z A12, dolna z A14
ws["A12"] = "Obaj porządkowi pełnią dyżur razem, przez całe zebranie. Trzecim porządkowym jest brat otwierający Salę."
grey = PatternFill("solid", fgColor=Color(indexed=22))
for r in (12, 13, 14):
    for c in "ABCD":
        cell = ws[f"{c}{r}"]
        cell.fill = grey
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(top=Side("medium") if r == 12 else Side(), bottom=Side("thin") if r == 14 else Side())
ws.merge_cells("A12:D14")
# czcionka spoza systemu: Excel na Macu rysuje zwykłą szeryfowo, pogrubioną bezszeryfowo
ws["D5"].font = Font(name="Czcionka tekstu podstawowego", size=11)
ws["D6"].font = Font(name="Czcionka tekstu podstawowego", size=11, bold=True)
# zwykła komórka z „Zawijaj tekst” (Układ jak w Excelu: zawija się, wiersz z pliku)
ws["B9"] = "opis zawijany w zwykłej komórce danych jak w Excelu"
ws["B9"].alignment = Alignment(wrap_text=True)
ws.row_dimensions[9].height = 30

# arkusz z tabelą Excela (Medium19: nagłówek w kolorze akcentu 4, szare paski) zapisany w 150%
t = wb.create_sheet("Tabela")
t.sheet_view.zoomScale = 150
t.append(["LP", "Osoba"])
for i in range(1, 7):
    t.append([i, "Osoba %03d" % i])
tab = Table(displayName="Lista", ref="A1:B7")
tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium19", showRowStripes=True)
t.add_table(tab)
t.column_dimensions["B"].width = 20
# wspólna krawędź zapisana w obu komórkach (dół B3 i góra B4 — medium): w Excelu jedna linia
t["B3"].border = Border(bottom=Side("medium"))
t["B4"].border = Border(top=Side("medium"))

wb.save(OUT)
print("zapisano", OUT)
