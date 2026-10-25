# -*- coding: utf-8 -*-
"""Builds the final delivery — ``Mazandasht_v04/``.

    python build4/make4.py            # full build (uses the harvested cache)
    python build4/make4.py --tiny     # smoke test with tiny logs

What happens
  1. 00_Master  — rebuilt, «باسکول دیگران» persons get names, tooltips added
  2. 01_Entry   — 50,000-line journal (10,000 weighbridge + 2,500 + 2,500),
                  Inv / Print / GRep sheets, tooltips, A4 everywhere
  3. 03/04/05/06 — Data mirrors grown to 50,000 lines, A4 print geometry
  4. 02_Bank    — rebuilt (18 operations, Payable report, tooltips)
  5. 08_Dash    — rebuilt on the 50,000-line feed + Prod analysis
"""
import json
import os
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(ROOT, "build2"))

import openpyxl                                                    # noqa: E402
from openpyxl.styles import Alignment, Font                         # noqa: E402

import kit                                                          # noqa: E402
import master                                                       # noqa: E402
import bank2                                                        # noqa: E402
import dash                                                         # noqa: E402
import links                                                        # noqa: E402
import unify                                                        # noqa: E402

import grow                                                         # noqa: E402
import entry4                                                       # noqa: E402
import mirror4                                                      # noqa: E402
import relink                                                       # noqa: E402
import bill                                                         # noqa: E402

V02 = os.path.join(ROOT, "Mazandasht_v02")
V03 = os.path.join(ROOT, "Mazandasht_v03")
OUT = os.path.join(ROOT, "Mazandasht_v04")
CACHE = os.path.join(ROOT, "build2", "feed_cache.json")

MIRRORS = ("03_Report.xlsx", "04_Inv.xlsx", "05_Bal.xlsx", "06_Buyers.xlsx")
COPIES = ("01_Entry.xlsx",) + MIRRORS + ("07_Demo.xlsx",)

PERSONS = [
    "احمد زاهدی", "احمد موسوی", "رضا کاظمی", "حسین نیک‌پور", "مهدی رستمی",
    "علی شریفی", "سعید اکبری", "مجید قاسمی", "پرویز احمدی", "کریم صادقی",
    "رحیم نوری", "بهرام یوسفی", "فرهاد مرادی", "مسعود تقوی", "جواد کریمی",
    "ناصر رحیمی", "بابک سلطانی", "وحید عباسی", "شخص ۱", "شخص ۲",
]


def bd_rows(master_path):
    """Bank legs for «باسکول دیگران» so the Payable report carries real movement."""
    wb = openpyxl.load_workbook(master_path, read_only=True)
    ws = wb["Acc"]
    people = []
    for r in range(10, 260):
        code = str(ws.cell(r, 1).value or "")
        if code.startswith("BD"):
            people.append((code, str(ws.cell(r, 4).value or "—")))
    wb.close()
    if len(people) < 3:
        return []
    pick = people[1:4] if len(people) > 3 else people[:3]
    out = [
        ("1405/07/04", "TB-902", "دریافت هزینه باسکول دیگران", "BN02", "",
         "کارت/پوز", "", "هزینه باسکول دیگران — دریافت از مشتری",
         "", 1_800_000, f"{pick[0][0]} – {pick[0][1]}", "هزینه باسکول دیگران"),
        ("1405/07/09", "TB-903", "دریافت هزینه باسکول دیگران", "BN01", "",
         "انتقال بانکی", "", "هزینه باسکول دیگران — واریز",
         "TR-9002", 2_400_000, f"{pick[1][0]} – {pick[1][1]}", "هزینه باسکول دیگران"),
        ("1405/07/11", "TB-904", "پرداخت به باسکول دیگران", "BN02", "",
         "کارت/پوز", "", "تسویه بدهی باسکول دیگران",
         "", 1_500_000, f"{pick[0][0]} – {pick[0][1]}", "تسویه باسکول دیگران"),
        ("1405/07/12", "TB-905", "پرداخت به باسکول دیگران", "BN01", "",
         "انتقال بانکی", "", "تسویه بخشی از بدهی باسکول دیگران",
         "TR-9003", 900_000, f"{pick[1][0]} – {pick[1][1]}", "تسویه باسکول دیگران"),
    ]
    return out


# --------------------------------------------------------------------- master
def persons(path):
    """Name the «باسکول دیگران» accounts and add tooltips to the mother file."""
    wb = openpyxl.load_workbook(path)
    ws = wb["Acc"]
    used = 0
    for r in range(10, 260):
        code = str(ws.cell(r, 1).value or "")
        if not code.startswith("BD"):
            continue
        name = ws.cell(r, 4).value
        if not name or "بدون نام" in str(name):
            pick = PERSONS[used % len(PERSONS)]
            ws.cell(r, 4).value = pick
            used += 1
        if not ws.cell(r, 2).value:
            ws.cell(r, 2).value = "باسکول دیگران"
        if not ws.cell(r, 9).value:
            ws.cell(r, 9).value = "طرف حساب"
        if not ws.cell(r, 10).value:
            ws.cell(r, 10).value = "بله"
        ws.cell(r, 6).value = ws.cell(r, 6).value or ""
    grey = Font(name="B Nazanin", size=9, italic=True, color="7A7A7A")
    note = ("این اشخاص، طرف حساب «باسکول دیگران» هستند: در فایل 02_Bank با نوع «دریافت/پرداخت باسکول دیگران» "
            "انتخاب می‌شوند و مانده بدهی هر کدام در شیت Payable همان فایل و در گزارش گروه‌ها (01_Entry ← GRep) "
            "دیده می‌شود. برای شخص جدید، نخستین کد خالی BD را پر کنید.")
    cell = ws.cell(9, 11, "راهنمای باسکول دیگران")
    cell.font = Font(name="B Nazanin", size=9, bold=True, color="FF14453D")
    ws.merge_cells(start_row=9, start_column=11, end_row=9, end_column=16)
    c = ws.cell(10, 11, note)
    c.font = grey
    c.alignment = Alignment(horizontal="right", vertical="top", wrap_text=True)
    ws.merge_cells(start_row=10, start_column=11, end_row=14, end_column=16)
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)
    return used


# ------------------------------------------------------------------- master+
def master_extras(path):
    """Adds the transaction-classification sheet (Cls) and the owner's current
    account (X021) to the mother file."""
    from openpyxl.styles import Border, PatternFill, Side
    wb = openpyxl.load_workbook(path)

    # ---- owner current account
    acc = wb["Acc"]
    used = [r for r in range(10, 300) if acc.cell(r, 1).value]
    codes = {str(acc.cell(r, 1).value) for r in used}
    if "X021" not in codes:
        row = (max(used) + 1) if used else 10
        vals = ["X021", "سرمایه و آورده صاحبان", "جاری شرکا", "حساب جاری رئیس شرکت",
                "—", "—", "—"]
        for i, v in enumerate(vals):
            cell = acc.cell(row, i + 1, v)
            cell.font = Font(name="B Nazanin", size=9)
            cell.alignment = Alignment(horizontal="center", vertical="center")

    # ---- classification sheet
    if "Cls" in wb.sheetnames:
        del wb["Cls"]
    ws = wb.create_sheet("Cls", 5)
    thin = Side(style="thin", color="DDD5C4")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    forest, white, sand = "14453D", "FFFFFF", "F7F1E6"
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:G1")
    c = ws["A1"]
    c.value = "   طبقه‌بندی تراکنش‌های بانکی — این پول برای چه کاری بود؟"
    c.font = Font(name="B Titr", size=14, bold=True, color=white)
    c.fill = PatternFill("solid", fgColor=forest)
    c.alignment = Alignment(horizontal="right", vertical="center")
    ws.row_dimensions[1].height = 30
    ws.merge_cells("A2:G2")
    c = ws["A2"]
    c.value = ("   هر ردیف یک زیرمجموعه است: خرید مرکبات، فروش مرکبات، قبض، حقوق، "
               "تعمیر و نگهداری دستگاه‌ها، و هزینه‌های شخصی رئیس. ستون «دامنه» مشخص می‌کند "
               "این پول شرکت است یا پول شخصی رئیس.")
    c.font = Font(name="B Nazanin", size=9, color="6B7A76")
    c.fill = PatternFill("solid", fgColor="FBF8F1")
    c.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 30

    heads = ["کد", "عنوان طبقه‌بندی", "دامنه", "گروه حساب", "توضیح", "فعال"]
    spans = [(1, 1), (2, 2), (3, 3), (4, 4), (5, 5), (6, 6)]
    for (a, b), h in zip(spans, heads):
        cell = ws.cell(8, a, h)
        cell.font = Font(name="B Nazanin", size=9.5, bold=True, color=white)
        cell.fill = PatternFill("solid", fgColor=forest)
        cell.border = box
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[8].height = 24

    ens = ["Code", "Classification", "Scope", "Group", "Note", "Active"]
    for j, h in enumerate(ens, start=1):
        cell = ws.cell(9, j, h)
        cell.font = Font(name="Tahoma", size=7.5, color="8B7E68")
        cell.fill = PatternFill("solid", fgColor="F3EADB")
        cell.border = box
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for i, (code, title, scope, group) in enumerate(bank2.CLASSES):
        r = 10 + i
        vals = [code, title, scope, group,
                "پول شخصی رئیس — به حساب جاری رئیس می‌رود" if scope != "شرکتی"
                else "هزینه/درآمد شرکت", "بله"]
        for j, v in enumerate(vals, start=1):
            cell = ws.cell(r, j, v)
            cell.font = Font(name="B Nazanin", size=9,
                             color="9A6B15" if scope != "شرکتی" else "2B2A26")
            cell.fill = PatternFill("solid",
                                    fgColor="FFF1D6" if scope != "شرکتی" else sand)
            cell.border = box
            cell.alignment = Alignment(horizontal="right" if j == 2 else "center",
                                       vertical="center")
    width = {"A": 9, "B": 36, "C": 13, "D": 24, "E": 42, "F": 8}
    for col, w in width.items():
        ws.column_dimensions[col].width = w
    ws.print_area = f"A1:F{9 + len(bank2.CLASSES)}"
    ws.print_title_rows = "8:9"
    ws.page_setup.paperSize = 9
    ws.page_setup.orientation = "portrait"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_options.horizontalCentered = True
    ws.oddHeader.left.text = "مازندشت — فایل مادر"
    ws.oddHeader.center.text = "MAZANDASHT · CLASSIFICATION"
    ws.oddHeader.right.text = "&A"
    ws.oddFooter.left.text = "&F"
    ws.oddFooter.center.text = "صفحه &P از &N"
    ws.oddFooter.right.text = "نسخه ۰۴ · ۱۴۰۵"
    for part in (ws.oddHeader.left, ws.oddHeader.center, ws.oddHeader.right,
                 ws.oddFooter.left, ws.oddFooter.center, ws.oddFooter.right):
        part.size = 8

    wb.calculation.fullCalcOnLoad = True
    wb.save(path)
    wb.close()
    return len(bank2.CLASSES)


# ---------------------------------------------------------------------- cache
def load_cache():
    with open(CACHE, encoding="utf-8") as fh:
        data = json.load(fh)
    f1 = {tuple(k.split("|")): v for k, v in data.get("F1", {}).items()}
    f2 = {tuple(k.split("|")): v for k, v in data.get("F2", {}).items()}
    return {"dates": data.get("dates") or [], "ops": data.get("ops") or [],
            "statement": data.get("statement") or [], "F1": f1, "F2": f2}


def groups_of(path):
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb["Grp"]
    out, seen = [], set()
    for r in range(10, 60):
        v = ws.cell(r, 1).value
        if v and v not in seen:
            seen.add(v)
            out.append(v)
    wb.close()
    return out[:11]


# ----------------------------------------------------------------------- main
def main(tiny=False):
    if tiny:
        grow.SCALE_LAST, grow.PAY_LAST, grow.CRATE_LAST = 29, 19, 19
        grow.BOOK_LAST = 129
        grow.DATA_LAST = 129
        grow.ENTRY_MAP = {1807: 129, 1808: 129, 1809: 129,
                          157: 19, 158: 19, 159: 19, 307: 29, 308: 29, 309: 29}
        entry4.INV_LAST = 29
        dash.BOOK_ROWS = 129
        dash.F1_LAST = 9 + 129
    else:
        dash.BOOK_ROWS = grow.BOOK_LAST - 9
        dash.F1_LAST = 9 + dash.BOOK_ROWS
    dash.POST_ROWS = 2 * bank2.NR
    dash.F2_LAST = 9 + dash.POST_ROWS

    os.makedirs(OUT, exist_ok=True)
    kit.ICONS = kit.make_icons()

    print("00_Master.xlsx")
    master.build(os.path.join(V02, "00_Master.xlsx"), os.path.join(OUT, "00_Master.xlsx"))
    n_cls = master_extras(os.path.join(OUT, "00_Master.xlsx"))
    print(f"  Cls sheet: {n_cls} classifications · owner account X021")
    n = persons(os.path.join(OUT, "00_Master.xlsx"))
    print(f"  named {n} «باسکول دیگران» accounts")
    print("   relink:", relink.polish(os.path.join(OUT, "00_Master.xlsx")))
    groups = groups_of(os.path.join(OUT, "00_Master.xlsx"))

    print("01_Entry.xlsx — growing the logs")
    for name in COPIES:
        src = os.path.join(V03, name)
        if os.path.exists(src):
            shutil.copy2(src, os.path.join(OUT, name))
    info = entry4.build(os.path.join(OUT, "01_Entry.xlsx"), groups)
    print(" ", info)
    bill.build(os.path.join(OUT, "01_Entry.xlsx"))
    print("   Bill sheet: auto invoice + payment block")
    print("   relink:", relink.polish(os.path.join(OUT, "01_Entry.xlsx")))

    if os.path.exists(os.path.join(OUT, "07_Demo.xlsx")):
        relink.polish(os.path.join(OUT, "07_Demo.xlsx"))

    print("03/04/05/06 — mirrors")
    for name in MIRRORS:
        p = os.path.join(OUT, name)
        sheets = mirror4.build(p)
        print(f"  {name}: {sheets} sheets polished")

    print("02_Bank.xlsx")
    vals = load_cache()
    demo = list(bank2.DEMO) + bd_rows(os.path.join(OUT, "00_Master.xlsx"))
    bank2.build(os.path.join(OUT, "00_Master.xlsx"), os.path.join(OUT, "02_Bank.xlsx"),
                demo=demo or None)
    print(f"  statement rows: {len(demo)}")
    unify.bal_ibank(os.path.join(OUT, "05_Bal.xlsx"))
    f2 = {ref: v for (c, ref), v in vals["F2"].items()}
    links.patch(os.path.join(OUT, "05_Bal.xlsx"),
                [{"file": "02_Bank.xlsx", "sheets": ["Post"], "values": {"Post": f2}}])

    print("08_Dash.xlsx")
    tmp = os.path.join(HERE, "_dash.xlsx")
    dash.build(os.path.join(OUT, "00_Master.xlsx"), os.path.join(OUT, "02_Bank.xlsx"),
               os.path.join(OUT, "01_Entry.xlsx"), tmp, vals)
    links.patch(tmp, [
        {"file": "01_Entry.xlsx", "sheets": ["Book"],
         "values": {"Book": {ref: v for (c, ref), v in vals["F1"].items()}}},
        {"file": "02_Bank.xlsx", "sheets": ["Post"], "values": {"Post": f2}},
    ])
    shutil.move(tmp, os.path.join(OUT, "08_Dash.xlsx"))
    print("  -> 08_Dash.xlsx", os.path.getsize(os.path.join(OUT, "08_Dash.xlsx")) // 1024,
          "KB")
    readme = os.path.join(HERE, "README_v04.md")
    if os.path.exists(readme):
        shutil.copy2(readme, os.path.join(OUT, "README.md"))
    print("done →", OUT)


if __name__ == "__main__":
    main("--tiny" in sys.argv)
