# -*- coding: utf-8 -*-
"""Numeric check of the two v05 report sheets (LBRep / GrwTot).

Run after ``valuecheck.py`` (which builds ``Mazandasht_check/``): this loads
``03_Report`` together with its ``01_Entry`` link target and prints the numbers
the user will see.
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
OUT = os.path.join(ROOT, "Mazandasht_check")

import formulas  # noqa: E402

ent = os.path.join(OUT, "01_Entry.xlsx")
rep = os.path.join(OUT, "03_Report.xlsx")
model = formulas.ExcelModel().loads(ent, rep).finish()
sol = model.calculate()


def get(sheet, ref):
    for key in (f"'[03_REPORT.XLSX]{sheet.upper()}'!{ref}",
                f"'[03_Report.xlsx]{sheet.upper()}'!{ref}"):
        if key in sol:
            v = sol[key]
            arr = getattr(v, "value", v)
            while hasattr(arr, "shape") and getattr(arr, "shape", None):
                arr = arr[0, 0] if arr.size else None
            return arr
    return "<miss>"


print("--- LBRep: one row per load buyer")
for r in range(10, 16):
    print(r, [get("LBRep", f"{c}{r}") for c in "ABCDEFGHIJ"])
print("total:", [get("LBRep", f"{c}30") for c in "DEFGHI"])

print("\n--- LBRep: stacked detail")
for r in range(34, 46):
    print(r, [get("LBRep", f"{c}{r}") for c in "ABCDEFGHIJ"])
print("total:", [get("LBRep", f"{c}154") for c in "FGHI"])

print("\n--- GrwTot: first growers")
for r in range(10, 16):
    print(r, [get("GrwTot", f"{c}{r}") for c in "ABCDEFGHIJK"])
print("total:", [get("GrwTot", f"{c}70") for c in "DEFGHI"])
print("\nKPI:", [get("GrwTot", f"{c}{r}") for r in (7,) for c in "ACEG"])
