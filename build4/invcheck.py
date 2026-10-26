# -*- coding: utf-8 -*-
""" Does the invoice sheet actually produce numbers?

``AGGREGATE`` is not supported by the ``formulas`` engine, so the invoice is
checked in two steps:

1. the row-selection itself is emulated in Python from the *literal* columns of
   ``Scale`` (date / operation / party) — exactly the condition the sheet uses;
2. those row numbers are written into a scratch copy (replacing only the
   ``AGGREGATE`` helper cells ``L11:L26`` and ``M33:M40``), and the whole sheet
   — lookups, tare, net weight, amount, totals, payment block, balance — is
   evaluated by the engine.

Run after ``make4.py --mid``:

    .venv/bin/python build4/invcheck.py [SL001] [C001]
"""
import os
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
OUT = os.path.join(ROOT, "Mazandasht_check")
SRC = os.path.join(OUT, "01_Entry.xlsx")
TMP = os.path.join(HERE, "_inv.xlsx")

import openpyxl  # noqa: E402

SCALE_LAST = 10009
PAY_LAST = 2509


def pick_scale(ws, party, op_word=""):
    """The rows *Bill* would show: same party, op filter, no partnership rows."""
    out = []
    for r in range(10, SCALE_LAST + 1):
        code = ws.cell(r, 7).value          # G — account / party code
        op = ws.cell(r, 5).value            # E — operation
        if not code or code != party or not op:
            continue
        if op_word and op_word not in str(op):
            continue
        if "شراکتی" in str(op):
            continue
        out.append(r)
    return out


def pick_pay(ws, party, kind):
    out = []
    for r in range(10, PAY_LAST + 1):
        if ws.cell(r, 6).value == party and ws.cell(r, 4).value == kind:
            out.append(r)
    return out


def patch(ws, col, first, rows, room):
    for i in range(room):
        ws[f"{col}{first + i}"] = rows[i] if i < len(rows) else ""


def main():
    sale = sys.argv[1] if len(sys.argv) > 1 else "SL001"
    buy = sys.argv[2] if len(sys.argv) > 2 else "C001"
    shutil.copy2(SRC, TMP)
    wb = openpyxl.load_workbook(TMP)
    sc, pay = wb["Scale"], wb["Pay"]
    plans = {"Bill": (sale, "دریافت", ""), "BillBuy": (buy, "پرداخت", "خرید")}
    plan_rows = {}
    for name, (party, kind, op_word) in plans.items():
        rows = pick_scale(sc, party, op_word)
        pays = pick_pay(pay, party, kind)
        plan_rows[name] = (party, rows, pays)
        patch(wb[name], "L", 11, rows, 16)
        patch(wb[name], "M", 33, pays, 8)
        wb[name]["O2"] = party
    wb.calculation.fullCalcOnLoad = True
    wb.save(TMP)
    wb.close()

    import formulas
    model = formulas.ExcelModel().loads(TMP).finish()
    sol = model.calculate()

    def get(sheet, ref):
        for key in (f"'[_INV.XLSX]{sheet.upper()}'!{ref}",
                    f"'[_inv.xlsx]{sheet.upper()}'!{ref}"):
            if key in sol:
                v = sol[key]
                arr = getattr(v, "value", v)
                while hasattr(arr, "shape") and getattr(arr, "shape", None):
                    arr = arr[0, 0] if arr.size else None
                return arr
        return "<miss>"

    for name, (party, rows, pays) in plan_rows.items():
        print(f"\n=== {name}   party {party}   scale rows {rows}   payrows {pays}")
        print("header:", get(name, "E7"), "|", get(name, "G7"), "|", get(name, "I7"))
        for r in range(11, 27):
            vals = [get(name, f"{c}{r}") for c in "ABCDEFGHIJ"]
            if any(v not in ("", None) for v in vals):
                print("  line", r - 10, vals)
        print("  totals F/G/H/J:", [get(name, f"{c}27") for c in "FGHJ"])
        print("  waste(kg/rial):", get(name, "H28"), get(name, "I28"),
              "| final kg/rial:", get(name, "H29"), get(name, "I29"))
        for r in range(33, 41):
            vals = [get(name, f"{c}{r}") for c in "ABCDEF"]
            if any(v not in ("", None) for v in vals):
                print("  pay", r - 32, vals)
        print("  pay total:", get(name, "D41"), "| balance:", get(name, "F41"))


if __name__ == "__main__":
    main()
