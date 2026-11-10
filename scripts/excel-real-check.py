#!/usr/bin/env python3
"""excel-real-check.py — pliki z aplikacji przez PRAWDZIWEGO Excela (Mac, Microsoft Excel 16+).

Dlaczego: walidator Open XML SDK i openpyxl sprawdzają schemat, ale nie to, co Excel zrobi
z plikiem. Tu Excel otwiera kopię każdego pliku, przelicza ją i zapisuje po swojemu, a my
porównujemy jego wersję z naszą. Jeśli Excel coś po cichu naprawił albo wyrzucił (tabelę,
zakres formatowania, formułę, format komórki), wyjdzie w różnicach. Plik, którego Excel nie
otworzy, = błąd.

  .venv/bin/python -I scripts/excel-real-check.py                 (wszystkie *.xlsx z folderu)
  .venv/bin/python -I scripts/excel-real-check.py plik.xlsx ...
  npm run test:excel   (najpierw zapisuje pliki z testów row-grow / col-grow, potem to)

Folder: /private/tmp/swb-excel-check. NIE ~/Documents — tam Office pyta o dostęp do każdego
pliku osobno. Otwieranie przez `open -a` (jak dwuklik w Finderze): bez okna zgody. Sterowanie
Excelem przez osascript wymaga tylko zgody „Automatyzacja” (raz). Excela nie zamykamy — po
teście zostaje uruchomiony bez otwartych plików testowych. Test lokalny (nie w test:serial):
wymaga Excela na tym komputerze.
"""
import os
import shutil
import subprocess
import sys
import time
import zipfile

import openpyxl

DIR = "/private/tmp/swb-excel-check"
SKIP_PARTS = ("docProps/", "xl/calcChain", "xl/printerSettings")


def osa(*lines, timeout=120):
    script = ["-e", f"with timeout of {timeout} seconds"]
    for line in lines:
        script += ["-e", line]
    script += ["-e", "end timeout"]
    r = subprocess.run(["osascript", *script], capture_output=True, text=True, timeout=timeout + 30)
    return r.returncode, (r.stdout or r.stderr).strip()


def workbook_names():
    code, out = osa('tell application "Microsoft Excel"', 'set n to ""',
                    'repeat with i from 1 to count of workbooks', 'set n to n & (name of workbook i) & linefeed', 'end repeat',
                    'return n', 'end tell', timeout=30)
    return set(out.splitlines()) if code == 0 else None


def excel_resave(src):
    """Kopia rt-<plik> otwarta, przeliczona i zapisana przez Excela. Zwraca (ścieżka|None, opis)."""
    name = os.path.basename(src)
    rt = os.path.join(DIR, "rt-" + name)
    shutil.copyfile(src, rt)
    subprocess.run(["open", "-a", "Microsoft Excel", rt], check=True)
    t0 = time.time()
    while time.time() - t0 < 90:
        names = workbook_names()
        if names is None:
            return None, "Excel nie odpowiada (ukryte okno? spójrz na ekran)"
        if os.path.basename(rt) in names:
            break
        time.sleep(1)
    else:
        return None, "Excel nie otworzył pliku w 90 s (odrzucony albo czeka na okno)"
    wb = os.path.basename(rt).replace('"', '\\"')
    code, out = osa('tell application "Microsoft Excel"', f'set wb to workbook "{wb}"', "calculate", "save wb",
                    "close wb saving no", "end tell")
    if code != 0:
        return None, f"zapis w Excelu nie wyszedł: {out[:200]}"
    return rt, f"otwarty i zapisany w {time.time() - t0:.0f} s"


def parts(p):
    with zipfile.ZipFile(p) as z:
        return {n for n in z.namelist() if not n.endswith("/") and not n.startswith(SKIP_PARTS)}


def sheet_facts(ws):
    return {
        "autofiltr": ws.auto_filter.ref,
        "formatowanie warunkowe": sorted(str(r.sqref) for r in ws.conditional_formatting),
        "listy wyboru": sorted(str(d.sqref) for d in ws.data_validations.dataValidation),
        "tabele": {t.displayName: (t.ref, [c.name for c in t.tableColumns]) for t in ws.tables.values()},
        "scalenia": sorted(str(m) for m in ws.merged_cells.ranges),
    }


def compare(ours, theirs):
    problems = []
    missing = sorted(parts(ours) - parts(theirs))
    if missing:
        problems.append(f"Excel usunął części: {missing}")
    a = openpyxl.load_workbook(ours)
    b = openpyxl.load_workbook(theirs)
    if a.sheetnames != b.sheetnames:
        problems.append(f"arkusze: {a.sheetnames} → {b.sheetnames}")
    for name in a.sheetnames:
        if name not in b.sheetnames:
            continue
        wa, wb = a[name], b[name]
        fa, fb = sheet_facts(wa), sheet_facts(wb)
        for k in fa:
            if fa[k] != fb[k]:
                problems.append(f"[{name}] {k}: nasz={fa[k]} | Excel={fb[k]}")
        bad_cells = []
        for row in wa.iter_rows():
            for x in row:
                if x.value is None and not x.has_style:
                    continue
                y = wb[x.coordinate]
                if (x.value, x.number_format) != (y.value, y.number_format):
                    bad_cells.append(f"{x.coordinate}: {x.value!r}/{x.number_format} → {y.value!r}/{y.number_format}")
        if bad_cells:
            problems.append(f"[{name}] komórki ({len(bad_cells)}): " + "; ".join(bad_cells[:5]))
    return problems


def main():
    os.makedirs(DIR, exist_ok=True)
    files = sys.argv[1:] or sorted(os.path.join(DIR, f) for f in os.listdir(DIR)
                                   if f.endswith(".xlsx") and not f.startswith(("rt-", "~$")))
    if not files:
        print(f"Brak plików w {DIR}")
        return 1
    if workbook_names() is None:
        print("❌ Excel nie odpowiada na polecenia (brak zgody „Automatyzacja” albo wisi okno)")
        return 1
    failed = 0
    for f in files:
        rt, note = excel_resave(f)
        if not rt:
            print(f"❌ {os.path.basename(f)} — {note}")
            failed += 1
            continue
        problems = compare(f, rt)
        print(f"{'✅' if not problems else '❌'} {os.path.basename(f)} — {note}")
        for pr in problems:
            print("   " + pr)
        failed += 1 if problems else 0
    print(f"{'✅' if not failed else '❌'} excel-real-check: {len(files) - failed}/{len(files)} zgodnych z Excelem")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
