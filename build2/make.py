# -*- coding: utf-8 -*-
"""Builds the whole Mazandasht v02 series from the v01 workbooks.

    python build2/make.py
"""
import os
import sys

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)

import xf      # noqa: E402
import entry   # noqa: E402

SRC = os.path.join(ROOT, "Mazandasht_System_v01")
OUT = os.path.join(ROOT, "Mazandasht_v02")
TMP = os.path.join(ROOT, "build2", "_tmp.xlsx")


def scale_dates():
    wb = openpyxl.load_workbook(os.path.join(SRC, "01_Register.xlsx"), read_only=True)
    ws = wb["Baskol"]
    seen = []
    for row in ws.iter_rows(min_row=10, max_row=309, min_col=2, max_col=2, values_only=True):
        v = row[0]
        if v and v not in seen:
            seen.append(str(v))
    wb.close()
    return seen


def main():
    os.makedirs(OUT, exist_ok=True)
    dates = scale_dates()
    print("scale dates:", len(dates), dates[:3], "…")

    print("01_Entry.xlsx")
    entry.build(os.path.join(SRC, "01_Register.xlsx"), TMP, dates)
    fixed = xf.transform_package(TMP, os.path.join(OUT, "01_Entry.xlsx"), "01_Entry.xlsx")
    os.remove(TMP)
    print(f"  -> 01_Entry.xlsx  {os.path.getsize(os.path.join(OUT,'01_Entry.xlsx'))//1024} KB")

    for old, new in xf.FILES.items():
        if new == "01_Entry.xlsx":
            continue
        src = os.path.join(SRC, old)
        dst = os.path.join(OUT, new)
        if not os.path.exists(src):
            print("  missing", old)
            continue
        f = xf.transform_package(src, dst, new)
        print(f"{new:<18} {os.path.getsize(dst)//1024:>4} KB   fixed {f}")


if __name__ == "__main__":
    main()
