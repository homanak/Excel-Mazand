# -*- coding: utf-8 -*-
"""Small openpyxl post-passes on the workbooks that carry no external-link cache.

    import relink; relink.polish(path, "00_Master.xlsx")

What it does
  1. repairs stale hyperlink targets left over from the v02/v03 names
  2. appends a «فایل‌های همراه» navigation strip on the Home sheet, so every
     workbook can jump to each of its eight siblings (relative paths, no drive)
  3. gives the locked Sync sheet real A4 print logic (it is a sheet too)
"""
import os

import re

import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.properties import PageSetupProperties

FILES = ["00_Master.xlsx", "01_Entry.xlsx", "02_Bank.xlsx", "03_Report.xlsx",
         "04_Inv.xlsx", "05_Bal.xlsx", "06_Buyers.xlsx", "07_Demo.xlsx",
         "08_Dash.xlsx"]
RENAME = {"02_Reports.xlsx": "02_Bank.xlsx", "01_Register.xlsx": "01_Entry.xlsx",
          "03_Reports.xlsx": "03_Report.xlsx", "05_Balance.xlsx": "05_Bal.xlsx",
          "06_Buyer.xlsx": "06_Buyers.xlsx", "08_Dashboard.xlsx": "08_Dash.xlsx"}
LABEL = {"00_Master.xlsx": "فایل مادر", "01_Entry.xlsx": "دفتر ثبت", "02_Bank.xlsx": "بانک",
         "03_Report.xlsx": "گزارش‌ها", "04_Inv.xlsx": "فاکتورها", "05_Bal.xlsx": "تراز",
         "06_Buyers.xlsx": "خریداران", "07_Demo.xlsx": "نمونه", "08_Dash.xlsx": "داشبورد"}

FOREST, GOLD, MINT, CREAM = "14453D", "C89B3C", "E4F0E8", "FBF8F1"
FA, EN = "B Nazanin", "Tahoma"
THIN = Side(style="thin", color="D8D2C2")


def _fix_links(wb):
    fixed = 0
    for ws in wb.worksheets:
        if ws.max_row > 3000:                     # never walk the 50k logs
            continue
        for row in ws.iter_rows():
            for c in row:
                h = c.hyperlink
                if h is None:
                    continue
                t = h.target or ""
                base = t.split("#")[0]
                if base in RENAME:
                    h.target = t.replace(base, RENAME[base])
                    fixed += 1
                elif base.endswith(".xlsx") and base not in FILES:
                    h.target = None
                    fixed += 1
    return fixed


def _nav(wb, name):
    fixed = _fix_links(wb)
    if "Home" not in wb.sheetnames:
        return fixed
    ws = wb["Home"]
    r = max(ws.max_row + 2, 20)
    while isinstance(ws.cell(r, 1), MergedCell):
        r += 1
    # avoid a second strip
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 400), max_col=2):
        for c in row:
            if isinstance(c.value, str) and c.value.strip().startswith("فایل‌های همراه"):
                return fixed
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
    head = ws.cell(r, 1, "     فایل‌های همراه — با یک کلیک باز می‌شوند (مسیر نسبی، همه در همین پوشه)")
    head.font = Font(name=FA, size=10.5, bold=True, color="FFFFFF")
    head.fill = PatternFill("solid", fgColor=FOREST)
    head.alignment = Alignment(horizontal="right", vertical="center")
    ws.row_dimensions[r].height = 24
    others = [f for f in FILES if f != name]
    for i, f in enumerate(others):
        c = ws.cell(r + 1 + i // 5, 1 + (i % 5) * 2)
        c.value = f"{LABEL.get(f, f)}  {f[:-5]}"
        c.hyperlink = f
        c.font = Font(name=FA, size=9, bold=True, color=FOREST)
        c.fill = PatternFill("solid", fgColor=MINT)
        c.border = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
        c.alignment = Alignment(horizontal="center", vertical="center")
    if others:
        ws.row_dimensions[r + 1].height = 22
        if len(others) > 5:
            ws.row_dimensions[r + 2].height = 22
    last = r + 1 + (0 if len(others) <= 5 else 1)
    area = ws.print_area
    if area:
        m = re.search(r"\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)", str(area))
        if m and int(m.group(4)) < last:
            ws.print_area = f"A1:{m.group(3)}{last}"
    else:
        ws.print_area = f"A1:J{last}"
    return fixed


def _sync_print(wb):
    if "Sync" not in wb.sheetnames:
        return False
    ws = wb["Sync"]
    ws.sheet_view.rightToLeft = True
    ws.page_setup.paperSize = 9
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.print_options.horizontalCentered = True
    ws.oddHeader.left.text = "مازندشت — سامانه حسابداری سورتینگ"
    ws.oddHeader.center.text = "MAZANDASHT · SYNC"
    ws.oddHeader.right.text = "&A"
    ws.oddFooter.left.text = "&F"
    ws.oddFooter.center.text = "صفحه &P از &N"
    ws.oddFooter.right.text = "نسخه ۰۴ · ۱۴۰۵"
    cols = max(ws.max_column, 1)
    rows = min(max(ws.max_row, 1), 300)
    ws.print_area = f"A1:{openpyxl.utils.get_column_letter(cols)}{rows}"
    return True


def polish(path, nav=True):
    wb = openpyxl.load_workbook(path)
    fixed = _nav(wb, os.path.basename(path)) if nav else 0
    synced = _sync_print(wb)
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)
    wb.close()
    return {"links_fixed": fixed, "sync_print": synced}
