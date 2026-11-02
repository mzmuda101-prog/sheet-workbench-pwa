"""Generuje scripts/fixtures/virt-merges.xlsx (test virt-merges-playwright.js).

Arkusz „Grafik”: 400 wierszy danych (nowy silnik przejmuje od 200) ze scaleniami jak w grafikach:
pionowe (2–4 wiersze) co 10 wierszy w kolumnie B, poziome (C:D) co 7 wierszy, blok 3×3 co 37 wierszy.
Arkusz „Opisy”: 400 wierszy z opisami różnej długości w wąskiej kolumnie (zawijanie → różne wysokości).
Uruchom: .venv/bin/python scripts/fixtures/make-virt-merges.py
"""
import os

from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "virt-merges.xlsx")
wb = Workbook()
ws = wb.active
ws.title = "Grafik"
ws.append(["Nr", "Grupa", "Osoba", "Zadanie", "Uwagi", "Status"])
for i in range(1, 401):
    ws.append([i, "Grupa %d" % (i // 10 + 1), "Osoba %03d" % (i % 37), "Zadanie %d" % i, "u%d" % i, "ok" if i % 3 else "brak"])
fill = PatternFill("solid", fgColor="DDEBF7")
merged = set()
def merge(ref, r0, r1, c0, c1, text):
    cells = {(r, c) for r in range(r0, r1 + 1) for c in range(c0, c1 + 1)}
    if cells & merged:
        return
    merged.update(cells)
    for r in range(r0, r1 + 1):
        for c in range(c0, c1 + 1):
            if (r, c) != (r0, c0):
                ws.cell(r, c).value = None
    ws.cell(r0, c0).value = text
    ws.cell(r0, c0).fill = fill
    ws.cell(r0, c0).alignment = Alignment(vertical="center", wrap_text=True)
    ws.merge_cells(ref)
for r in range(2, 402, 10):  # pionowe w B, wysokość 2–4
    h = 2 + (r // 10) % 3
    if r + h - 1 <= 401:
        merge(f"B{r}:B{r + h - 1}", r, r + h - 1, 2, 2, "Grupa scalona %d" % r)
for r in range(5, 402, 7):  # poziome C:D
    merge(f"C{r}:D{r}", r, r, 3, 4, "Poziomo %d" % r)
for r in range(20, 400, 37):  # blok 3×3 E:G… (E:F + 3 wiersze)
    merge(f"E{r}:F{r + 2}", r, r + 2, 5, 6, "Blok %d" % r)
o = wb.create_sheet("Opisy")
o.append(["Nr", "Opis", "Kwota"])
words = "zebranie porządkowi dyżur sala brat otwierający nagłośnienie mikrofony scena grupa".split()
for i in range(1, 401):
    n = 2 + (i * 7) % 23  # 2–24 słów → 1–6 linii po zawinięciu
    o.append([i, " ".join(words[(i + k) % len(words)] for k in range(n)), i * 3])
o.column_dimensions["B"].width = 22
wb.save(OUT)
print("zapisano", OUT)
