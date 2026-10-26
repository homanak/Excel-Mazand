# -*- coding: utf-8 -*-
"""End-to-end *value* check of the partnership maths.

Builds a small copy (--mid: 60 weighbridge rows, 400 journal lines) into
``Mazandasht_check/``, then evaluates it with the ``formulas`` engine so the
share columns and the partner ledger are checked as numbers, not as text.
"""
import json
import os
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(ROOT, "build2"))

import make4  # noqa: E402


def main():
    out = os.path.join(ROOT, "Mazandasht_check")
    shutil.rmtree(out, ignore_errors=True)
    make4.OUT = out
    make4.main(mid=True)

    import formulas
    path = os.path.join(out, "01_Entry.xlsx")
    model = formulas.ExcelModel().loads(path).finish()
    sol = model.calculate()

    def get(sheet, ref):
        for key in (f"'[01_Entry.xlsx]{sheet.upper()}'!{ref}",
                    f"'[01_ENTRY.XLSX]{sheet.upper()}'!{ref}"):
            if key in sol:
                v = sol[key]
                arr = getattr(v, "value", v)
                while hasattr(arr, "shape") and getattr(arr, "shape", None):
                    arr = arr[0, 0] if arr.size else None
                return arr
        return "<miss>"

    print("\n--- Scale (partnership rows)")
    for r in range(28, 34):
        print(r, "op=", repr(get("Scale", f"E{r}")),
              "kg=", get("Scale", f"N{r}"),
              "amount=", get("Scale", f"AJ{r}"),
              "partner=", get("Scale", f"AQ{r}"),
              "our%=", get("Scale", f"AR{r}"),
              "our_rial=", get("Scale", f"AS{r}"),
              "partner_rial=", get("Scale", f"AT{r}"),
              "our_kg=", get("Scale", f"AU{r}"),
              "partner_kg=", get("Scale", f"AV{r}"),
              "leg_dr=", get("Scale", f"V{r}"), "leg_cr=", get("Scale", f"W{r}"),
              "bal=", get("Scale", f"AI{r}"), "ctl=", repr(get("Scale", f"AW{r}")))

    print("\n--- Shrk ledger")
    for r in range(14, 17):
        print(r, get("Shrk", f"A{r}"), get("Shrk", f"B{r}"),
              "ratio=", get("Shrk", f"C{r}"), "our=", get("Shrk", f"D{r}"),
              "in_kg=", get("Shrk", f"E{r}"), "in_val=", get("Shrk", f"F{r}"),
              "our_buy=", get("Shrk", f"G{r}"), "partner_buy=", get("Shrk", f"H{r}"),
              "out_kg=", get("Shrk", f"I{r}"), "our_sale=", get("Shrk", f"K{r}"),
              "paid=", get("Shrk", f"L{r}"), "bal=", get("Shrk", f"M{r}"),
              "st=", repr(get("Shrk", f"N{r}")))
    print("total:", [get("Shrk", f"{c}34") for c in "EFGHIJKLM"])

    print("\n--- detail block")
    for r in range(38, 44):
        print(r, [get("Shrk", f"{c}{r}") for c in "ABCDEFGHIJK"])

    print("\n--- split & fee (Pay rows)")
    for r in range(22, 27):
        print(r, "doc=", get("Pay", f"C{r}"), "type=", get("Pay", f"D{r}"),
              "acct=", get("Pay", f"F{r}"), "amount=", get("Pay", f"K{r}"),
              "| fee=", get("Pay", f"AH{r}"), get("Pay", f"AI{r}"),
              "splits=", (get("Pay", f"AJ{r}"), get("Pay", f"AK{r}"),
                          get("Pay", f"AL{r}"), get("Pay", f"AM{r}")),
              "| sum=", get("Pay", f"AQ{r}"), "net1=", get("Pay", f"AT{r}"),
              "net2=", get("Pay", f"AU{r}"), "ctl=", repr(get("Pay", f"AP{r}")))

    print("\n--- journal legs of PV-104 / PV-105")
    base = 10 + (make4.grow.SCALE_LAST - 9) * 4
    for prow in (25, 26):
        r0 = base + (prow - 10) * 6
        print("Pay", prow, "-> Book", r0, "…", r0 + 5)
        for r in range(r0, r0 + 6):
            print("   ", r, [get("Book", f"{c}{r}") for c in ("A", "D", "L", "M")])
        dr = sum((get("Book", f"L{r}") or 0) for r in range(r0, r0 + 6)
                 if isinstance(get("Book", f"L{r}"), (int, float)))
        cr = sum((get("Book", f"M{r}") or 0) for r in range(r0, r0 + 6)
                 if isinstance(get("Book", f"M{r}"), (int, float)))
        print("    legs debit", dr, "credit", cr, "diff", round(dr - cr, 2))

    print("\n--- Stmt (C001) — loads")
    for r in range(13, 20):
        print(r, [get("Stmt", f"{c}{r}") for c in "ABCDEFGHIJ"])
    print("Stmt counters:", get("Stmt", "M2"), get("Stmt", "M3"),
          "| idx:", [get("Stmt", f"M{r}") for r in (12, 13, 42, 45, 46, 53)])
    print("Stmt totals:", [get("Stmt", f"E{r}") for r in (54, 55, 56)],
          "| bal:", get("Stmt", "B8"), "| who:", get("Stmt", "D9"))

    print("\n--- Ledger (C001)")
    for r in range(12, 20):
        print(r, [get("Ledger", f"{c}{r}") for c in "ABCDEFGH"])
    print("Ledger totals:", [get("Ledger", f"{c}162") for c in "AFGH"],
          [get("Ledger", f"{c}163") for c in "FGH"], "| kpi:", get("Ledger", "A6"))
    print("Ledger page 40-50:", [(get("Ledger", f"A{r}"), get("Ledger", f"D{r}"))
                                 for r in range(40, 48)])

    print("\n--- journal totals")
    tot = make4.grow.BOOK_LAST + 1
    print("Book total row", tot, "| debit", get("Book", f"L{tot}"),
          "credit", get("Book", f"M{tot}"), "check", repr(get("Book", f"P{tot}")))
    print("Book total row cells:", [get("Book", f"{c}{tot}") for c in "ABCDEFGHIJKLMN"])


if __name__ == "__main__":
    main()
