# -*- coding: utf-8 -*-
"""08_Dash.xlsx — the dashboard file (v03).

Reads live from 01_Entry.xlsx (Book) and 02_Bank.xlsx (Post) through external
links, mirrors them into two hidden feed sheets and builds every chart and KPI
from local formulas — so the workbook recalculates instantly when opened.
"""
import json
import os
import re
import shutil
import zipfile

import openpyxl
import xlsxwriter

from kit import (Kit, FOREST, FOREST_2, SAGE, SAGE_L, MINT, CREAM, SAND, SAND_2,
                 GOLD, GOLD_L, AMBER, TERRA, BLUSH, SKY, SKY_D, LAVENDER, ROSE,
                 PEACH, GREY_BG, LINE, LINE_SOFT, INK, MUTED, WHITE, RED, GREEN,
                 BLUE, FA, FA_TITLE, EN, INT, INT_RED, PCT)

FILE = "08_Dash.xlsx"
ENTRY = "01_Entry.xlsx"
BANK = "02_Bank.xlsx"

BOOK_ROWS = 1799     # Book!A10..A1808
POST_ROWS = 280      # Post!A13..A292

BOOK_COLS = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O"]
BOOK_SRC = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O"]
POST_SRC = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"]

TEXT_COLS_F1 = {0, 1, 2, 3, 4, 5, 9, 13, 14}       # date, doc, source, account, desc, product, op, group
TEXT_COLS_F2 = {0, 1, 2, 3, 4, 5, 6, 7, 11}


# ---------------------------------------------------------------- feed sheets

# ====================================================================== Prod
def sheet_prod(wb, k, values):
    """Products: weight, money and share — built from the journal feed."""
    ws = wb.add_worksheet("Prod")
    prods = {}
    for (col, ref), v in (values or {}).items():
        if col != "F" or not isinstance(v, str) or not v.strip():
            continue
        row = int("".join(ch for ch in ref if ch.isdigit()))
        amount = float(values.get(("K", f"K{row}")) or 0)
        weight = float(values.get(("H", f"H{row}")) or 0)
        cur = prods.setdefault(v.strip(), [0.0, 0.0, 0])
        cur[0] += amount
        cur[1] += weight
        cur[2] += 1
    top = sorted(prods.items(), key=lambda kv: -kv[1][0])[:12] or [("—", [0, 0, 0])]
    n = len(top)
    last = 6
    k.setup(ws, last, 12 + n + 8, landscape=True, repeat=(11, 11), freeze=(13, 2))
    k.brand(ws, last, "تحلیل محصول",
            "وزن، مبلغ و سهم هر محصول  ·  از دفتر ثبت روزانه", icon="crate")
    k.kpis(ws, 6, [
        ("تعداد محصول", f"=COUNTA($A$13:$A${12+n})", INT, FOREST, "list"),
        ("جمع وزن (kg)", f"=SUM($C$13:$C${12+n})", INT, GREEN, "scale"),
        ("جمع مبلغ", f"=SUM($D$13:$D${12+n})", INT, GOLD, "money"),
        ("میانگین فی", f"=IFERROR($D$8/$C$8,0)", INT, TERRA, "gauge"),
    ], last)
    heads = [("محصول", "Product", 22), ("سطر", "Lines", 8), ("وزن (kg)", "Weight", 15),
             ("مبلغ", "Amount", 18), ("فی میانگین", "Avg price", 13),
             ("سهم مبلغ", "Share", 10), ("نمای نسبی", "Visual", 24)]
    for i, (fa, en, w) in enumerate(heads):
        ws.set_column(i, i, w)
        ws.write(11, i, fa, k.f(font_name=FA, font_size=9, bold=True, font_color=WHITE,
                                bg_color=FOREST, border=1, border_color=FOREST,
                                align="center", valign="vcenter"))
    ws.set_row(11, 28)
    for i, (name, (amount, weight, lines)) in enumerate(top):
        r = 12 + i
        ws.set_row(r, 20)
        ws.write_string(r, 0, name, k.f(font_name=FA, font_size=9, bold=True,
                                        font_color=FOREST, bg_color=SAND_2, border=1,
                                        border_color=LINE, align="right"))
        ws.write_formula(r, 1, f'=COUNTIF(F1!$F${F1_FIRST}:$F${F1_LAST},$A{r+1})',
                         k.f(font_name=EN, font_size=9, bg_color=WHITE, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(r, 2, f'=SUMIF(F1!$F${F1_FIRST}:$F${F1_LAST},$A{r+1},'
                               f'F1!$H${F1_FIRST}:$H${F1_LAST})',
                         k.f(font_name=EN, font_size=9, bg_color=MINT, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(r, 3, f'=SUMIF(F1!$F${F1_FIRST}:$F${F1_LAST},$A{r+1},'
                               f'F1!$K${F1_FIRST}:$K${F1_LAST})',
                         k.f(font_name=EN, font_size=9, bold=True, bg_color=GOLD_L, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(r, 4, f'=IFERROR($D{r+1}/$C{r+1},0)',
                         k.f(font_name=EN, font_size=9, bg_color=WHITE, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(r, 5, f'=IFERROR($D{r+1}/SUM($D$13:$D${12+n}),0)',
                         k.f(font_name=EN, font_size=8.5, bg_color=WHITE, border=1,
                             border_color=LINE_SOFT, align="center", num_format=PCT))
        ws.write_formula(r, 6, f'=REPT("█",ROUND(IF(MAX($D$13:$D${12+n})=0,0,'
                               f'$D{r+1}/MAX($D$13:$D${12+n})*18),0))',
                         k.f(font_name=EN, font_size=9, font_color=SAGE, bg_color=WHITE,
                             border=1, border_color=LINE_SOFT, align="right"))
    tot = 12 + n
    ws.merge_range(tot, 0, tot, 1, "جمع", k.f(font_name=FA, font_size=10, bold=True,
                                              font_color=WHITE, bg_color=GOLD, border=1))
    for c in (2, 3):
        ws.write_formula(tot, c, f"=SUM({chr(65+c)}13:{chr(65+c)}{tot})",
                         k.f(font_name=EN, font_size=10, bold=True, font_color=WHITE,
                             bg_color=GOLD, border=1, num_format=INT))
    ch = wb.add_chart({"type": "column"})
    ch.add_series({"name": "مبلغ", "categories": ["Prod", 12, 0, 11 + n, 0],
                   "values": ["Prod", 12, 3, 11 + n, 3], "fill": {"color": GOLD},
                   "gap": 45, "data_labels": {"value": True, "num_format": "#,##0",
                                              "font": {"size": 6.5}}})
    ch.set_title({"name": "مبلغ فروش/خرید هر محصول", "name_font": {"name": FA, "size": 11}})
    ch.set_legend({"none": True})
    ch.set_plotarea({"fill": {"color": "#FCFAF4"}})
    ch.set_size({"width": 560, "height": 250})
    ws.insert_chart(f"A{tot+3}", ch, {"x_offset": 8, "y_offset": 6})
    k.note(ws, tot + 20, last,
           "این شیت از ستون «محصول» دفتر روزنامه ساخته می‌شود؛ هر محصولی که در فایل مادر تعریف شده و "
           "در ثبت روزانه استفاده شده، خودکار اینجا می‌آید.", bg=SAND, icon="crate")
    ws.print_area(0, 0, tot + 21, last)
    return ws


def sheet_f1(wb, k, values):
    ws = wb.add_worksheet("F1")
    ws.hide()
    k.setup(ws, 14, 1808, landscape=True)
    for i, h in enumerate(["تاریخ", "سند", "منبع", "حساب", "شرح", "محصول", "سبد",
                           "وزن باسکول", "وزن خالص", "فی", "مبلغ", "بدهکار",
                           "بستانکار", "عملیات", "گروه", "ماه"]):
        ws.write(8, i, h, k.f(font_name=FA, font_size=8, bold=True, font_color=WHITE,
                              bg_color=FOREST, align="center"))
    for i in range(BOOK_ROWS):
        r = 9 + i
        src_row = 10 + i
        for c, col in enumerate(BOOK_COLS):
            ref = f"{col}{src_row}"
            val = values.get((col, ref), 0)
            if c in TEXT_COLS_F1:
                if not val:
                    continue
                ws.write_formula(r, c, f"='[{ENTRY}]Book'!{ref}",
                                 k.f(font_name=FA, font_size=8, align="right"), str(val))
            else:
                ws.write_formula(r, c, f"='[{ENTRY}]Book'!{ref}",
                                 k.f(font_name=EN, font_size=8, num_format=INT,
                                     align="center"),
                                 float(val) if isinstance(val, (int, float)) else 0)
        day = values.get(("A", f"A{src_row}")) or ""
        mon = int(day[5:7]) if isinstance(day, str) and len(day) >= 7 and day[5:7].isdigit() else ""
        ws.write_formula(r, 15,
                         f'=IF($A{src_row}="","",IFERROR(VALUE(MID($A{src_row},6,2)),""))',
                         k.f(font_name=EN, font_size=8, num_format="00", align="center"),
                         mon)
    return ws


def sheet_f2(wb, k, values):
    ws = wb.add_worksheet("F2")
    ws.hide()
    k.setup(ws, 11, 292, landscape=True)
    for i, h in enumerate(["تاریخ", "سند", "منبع", "حساب", "نام", "گروه", "عملیات",
                           "بانک", "بدهکار", "بستانکار", "ماه", "شرح"]):
        ws.write(8, i, h, k.f(font_name=FA, font_size=8, bold=True, font_color=WHITE,
                              bg_color=FOREST_2, align="center"))
    for i in range(POST_ROWS):
        r = 9 + i
        ref_row = 13 + i
        for c, col in enumerate(POST_SRC):
            ref = f"{col}{ref_row}"
            val = values.get((col, ref), 0)
            if c in TEXT_COLS_F2:
                if not val:
                    continue
                ws.write_formula(r, c, f"='[{BANK}]Post'!{ref}",
                                 k.f(font_name=FA, font_size=8, align="right"), str(val))
            else:
                ws.write_formula(r, c, f"='[{BANK}]Post'!{ref}",
                                 k.f(font_name=EN, font_size=8, num_format=INT,
                                     align="center"),
                                 float(val) if isinstance(val, (int, float)) else 0)
    return ws


# -------------------------------------------------------------------- tables
F1_FIRST, F1_LAST = 10, 9 + BOOK_ROWS          # excel rows
F2_FIRST, F2_LAST = 10, 9 + POST_ROWS


def sheet_time(wb, k):
    ws = wb.add_worksheet("Time")
    last = 6
    k.setup(ws, last, 42, landscape=True, repeat=(11, 11), freeze=(13, 1))
    k.brand(ws, last, "روند ماهانه عملیات",
            "وزن، مبلغ و گردش هر ماه شمسی  ·  از دفتر ثبت روزانه", icon="chart")
    k.kpis(ws, 6, [
        ("کل وزن (kg)", "=SUM($C$13:$C$24)", INT, GREEN, "scale"),
        ("کل مبلغ", "=SUM($D$13:$D$24)", INT, GOLD, "money"),
        ("کل بدهکار", "=SUM($E$13:$E$24)", INT, BLUE, "transfer"),
        ("کل بستانکار", "=SUM($F$13:$F$24)", INT, TERRA, "transfer"),
    ], last)
    heads = [("ماه", "Month", 10), ("تعداد گردش", "Count", 12), ("وزن (kg)", "Weight", 16),
             ("مبلغ", "Amount", 18), ("بدهکار", "Debit", 16), ("بستانکار", "Credit", 16),
             ("خالص", "Net", 16)]
    for i, (fa, en, w) in enumerate(heads):
        ws.set_column(i, i, w)
        ws.write(11, i, f"{fa}\n{en}",
                 k.f(font_name=FA, font_size=9, bold=True, font_color=WHITE,
                    bg_color=FOREST, border=1, border_color=FOREST, align="center",
                    valign="vcenter", text_wrap=True))
    ws.set_row(11, 32)
    f_mono = "F1!$P${0}:$P${1}".format(F1_FIRST, F1_LAST)
    for m in range(12):
        r = 12 + m
        ws.set_row(r, 21)
        ws.write_number(r, 0, m + 1, k.f(font_name=EN, font_size=9, bold=True,
                                         font_color=FOREST, bg_color=SAND_2, border=1,
                                         border_color=LINE, align="center", num_format="00"))
        ws.write_formula(r, 1, f'=COUNTIF({f_mono},$A{r+1})',
                         k.f(font_name=EN, font_size=9, bg_color=WHITE, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(r, 2, f'=SUMIF({f_mono},$A{r+1},F1!$H${F1_FIRST}:$H${F1_LAST})',
                         k.f(font_name=EN, font_size=9, bg_color=MINT, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(r, 3, f'=SUMIF({f_mono},$A{r+1},F1!$K${F1_FIRST}:$K${F1_LAST})',
                         k.f(font_name=EN, font_size=9, bg_color=GOLD_L, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(r, 4, f'=SUMIF({f_mono},$A{r+1},F1!$L${F1_FIRST}:$L${F1_LAST})',
                         k.f(font_name=EN, font_size=9, bg_color=WHITE, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(r, 5, f'=SUMIF({f_mono},$A{r+1},F1!$M${F1_FIRST}:$M${F1_LAST})',
                         k.f(font_name=EN, font_size=9, bg_color=WHITE, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(r, 6, f'=$E{r+1}-$F{r+1}',
                         k.f(font_name=EN, font_size=9, bold=True, bg_color=CREAM,
                             border=1, border_color=LINE_SOFT, align="center",
                             num_format=INT_RED))
    ch = wb.add_chart({"type": "column"})
    ch.add_series({"name": "وزن (kg)", "categories": ["Time", 12, 0, 23, 0],
                   "values": ["Time", 12, 2, 23, 2], "fill": {"color": SAGE}, "gap": 60})
    ln = wb.add_chart({"type": "line"})
    ln.add_series({"name": "مبلغ", "categories": ["Time", 12, 0, 23, 0],
                   "values": ["Time", 12, 3, 23, 3], "y2_axis": True,
                   "line": {"color": GOLD, "width": 2.4},
                   "marker": {"type": "circle", "size": 5, "fill": {"color": WHITE},
                              "border": {"color": GOLD}}})
    ch.combine(ln)
    ch.set_title({"name": "وزن و مبلغ ماهانه", "name_font": {"name": FA, "size": 11}})
    ch.set_legend({"position": "bottom", "font": {"name": FA, "size": 8}})
    ch.set_plotarea({"fill": {"color": "#FCFAF4"}})
    ch.set_size({"width": 640, "height": 290})
    ws.insert_chart("A27", ch, {"x_offset": 8, "y_offset": 6})
    return ws


def sheet_ops(wb, k, ops):
    ws = wb.add_worksheet("Ops")
    last = 6
    k.setup(ws, last, 32, landscape=True, repeat=(11, 11), freeze=(13, 1))
    k.brand(ws, last, "تحلیل عملیات",
            "کدام نوع عملیات چقدر وزن و مبلغ جابه‌جا کرده است؟", icon="gauge")
    k.kpis(ws, 6, [
        ("انواع عملیات", f"={len(ops)}", INT, FOREST, "list"),
        ("بیشترین وزن", "=MAX($C$13:$C$28)", INT, GREEN, "scale"),
        ("بیشترین مبلغ", "=MAX($E$13:$E$28)", INT, GOLD, "money"),
        ("ناترازها", "=COUNTIF(F1!$L$10:$L$1808,\">0\")-COUNTIF(F1!$M$10:$M$1808,\">0\")", INT, TERRA, "warn"),
    ], last)
    heads = [("نوع عملیات", "Operation", 30), ("تعداد", "Count", 10),
             ("وزن (kg)", "Weight", 16), ("مبلغ", "Amount", 18),
             ("بدهکار", "Debit", 16), ("بستانکار", "Credit", 16), ("سهم مبلغ", "Share", 10)]
    for i, (fa, en, w) in enumerate(heads):
        ws.set_column(i, i, w)
        ws.write(11, i, f"{fa}\n{en}",
                 k.f(font_name=FA, font_size=9, bold=True, font_color=WHITE,
                    bg_color=FOREST, border=1, border_color=FOREST, align="center",
                    text_wrap=True))
    ws.set_row(11, 30)
    for i, op in enumerate(ops[:16]):
        r = 12 + i
        ws.write_string(r, 0, op, k.f(font_name=FA, font_size=9, bg_color=SAND_2,
                                      border=1, border_color=LINE, align="right"))
        ws.write_formula(r, 1, f'=COUNTIF(F1!$N${F1_FIRST}:$N${F1_LAST},$A{r+1})',
                         k.f(font_name=EN, font_size=9, bg_color=WHITE, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(r, 2, f'=SUMIF(F1!$N${F1_FIRST}:$N${F1_LAST},$A{r+1},F1!$I${F1_FIRST}:$I${F1_LAST})',
                         k.f(font_name=EN, font_size=9, bg_color=MINT, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(r, 3, f'=SUMIF(F1!$N${F1_FIRST}:$N${F1_LAST},$A{r+1},F1!$K${F1_FIRST}:$K${F1_LAST})',
                         k.f(font_name=EN, font_size=9, bg_color=WHITE, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(r, 4, f'=SUMIF(F1!$N${F1_FIRST}:$N${F1_LAST},$A{r+1},F1!$L${F1_FIRST}:$L${F1_LAST})',
                         k.f(font_name=EN, font_size=9, bg_color=WHITE, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(r, 5, f'=SUMIF(F1!$N${F1_FIRST}:$N${F1_LAST},$A{r+1},F1!$M${F1_FIRST}:$M${F1_LAST})',
                         k.f(font_name=EN, font_size=9, bg_color=WHITE, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(r, 6, f'=IFERROR($D{r+1}/SUM($D$13:$D$28),0)',
                         k.f(font_name=EN, font_size=8.5, bg_color=WHITE, border=1,
                             border_color=LINE_SOFT, align="center", num_format=PCT))
    ch = wb.add_chart({"type": "bar"})
    ch.add_series({"name": "وزن", "categories": ["Ops", 11, 0, 26, 0],
                   "values": ["Ops", 11, 2, 26, 2], "fill": {"color": SAGE},
                   "data_labels": {"value": True, "num_format": "#,##0", "font": {"size": 7}}})
    ch.set_title({"name": "وزن هر نوع عملیات (kg)", "name_font": {"name": FA, "size": 11}})
    ch.set_legend({"none": True})
    ch.set_plotarea({"fill": {"color": "#FCFAF4"}})
    ch.set_size({"width": 520, "height": 300})
    ws.insert_chart("A30", ch, {"x_offset": 8, "y_offset": 6})
    return ws


def sheet_pos(wb, k, dates):
    ws = wb.add_worksheet("Pos")
    last = 7
    k.setup(ws, last, 12 + 45 + 4, landscape=True, repeat=(11, 11), freeze=(13, 2))
    k.brand(ws, last, "کنترل پز روزانه و واریز آن به بانک",
            "پز ثبت‌شده در دفتر روزنامه  ↔  واریزی همان روز به حساب بانکی  ↔  سایر واریزها", icon="card")
    k.kpis(ws, 6, [
        ("پز دفتر روزنامه", "=SUM($C$13:$C$52)", INT, GOLD, "pos"),
        ("واریز پز به بانک", "=SUM($D$13:$D$52)", INT, GREEN, "bank"),
        ("سایر واریزها", "=SUM($E$13:$E$52)", INT, SAGE, "card"),
        ("اختلاف", "=SUM($F$13:$F$52)", INT_RED, RED, "warn"),
    ], last)
    heads = [("تاریخ", "Date", 13), ("تاریخ واریز", "Deposit", 13),
             ("پز دفتر", "Journal POS", 17), ("واریز پز به بانک", "POS bank", 17),
             ("سایر واریزها", "Other", 17), ("اختلاف", "Diff", 15),
             ("وضعیت", "Status", 13)]
    for i, (fa, en, w) in enumerate(heads):
        ws.set_column(i, i, w)
        ws.write(11, i, f"{fa}\n{en}",
                 k.f(font_name=FA, font_size=9, bold=True, font_color=WHITE,
                    bg_color=FOREST, border=1, border_color=FOREST, align="center",
                    text_wrap=True, valign="vcenter"))
    ws.set_row(11, 30)
    for i in range(40):
        r = 12 + i
        d = dates[i] if i < len(dates) else ""
        ws.set_row(r, 20)
        ws.write_string(r, 0, d, k.f(font_name=EN, font_size=9, bg_color="#FFF4D6",
                                     border=1, border_color=GOLD_L, align="center"))
        ws.write_blank(r, 1, None, k.f(font_name=EN, font_size=9, bg_color="#FFF4D6",
                                       border=1, border_color=GOLD_L, align="center"))
        ws.write_formula(
            r, 2,
            f'=IF($A{r+1}="","",SUMIFS(F1!$L${F1_FIRST}:$L${F1_LAST},F1!$A${F1_FIRST}:$A${F1_LAST},'
            f'$A{r+1},F1!$D${F1_FIRST}:$D${F1_LAST},"X011",'
            f'F1!$N${F1_FIRST}:$N${F1_LAST},"پوز"))',
            k.f(font_name=EN, font_size=9, bg_color=GOLD_L, border=1,
                border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(
            r, 3,
            f'=IF($A{r+1}="","",SUMIFS(F2!$I${F2_FIRST}:$I${F2_LAST},F2!$A${F2_FIRST}:$A${F2_LAST},'
            f'IF($B{r+1}="",$A{r+1},$B{r+1}),F2!$D${F2_FIRST}:$D${F2_LAST},"BN*",'
            f'F2!$G${F2_FIRST}:$G${F2_LAST},"واریز پز باسکول (۲۰۰۶)"))',
            k.f(font_name=EN, font_size=9, bg_color=MINT, border=1,
                border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(
            r, 4,
            f'=IF($A{r+1}="","",SUMIFS(F2!$I${F2_FIRST}:$I${F2_LAST},F2!$A${F2_FIRST}:$A${F2_LAST},'
            f'IF($B{r+1}="",$A{r+1},$B{r+1}),F2!$D${F2_FIRST}:$D${F2_LAST},"BN*",'
            f'F2!$G${F2_FIRST}:$G${F2_LAST},"تسویه دریافت (واسط)"))',
            k.f(font_name=EN, font_size=9, bg_color=WHITE, border=1,
                border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(r, 5, f'=IF($A{r+1}="","",$C{r+1}-$D{r+1})',
                         k.f(font_name=EN, font_size=9, bold=True, bg_color=CREAM,
                             border=1, border_color=LINE_SOFT, align="center",
                             num_format=INT_RED))
        ws.write_formula(
            r, 6, f'=IF($A{r+1}="","",IF(ABS($F{r+1})<=1000,"✓ تراز","⚠ بررسی شود"))',
            k.f(font_name=FA, font_size=9, bold=True, bg_color=CREAM, border=1,
                border_color=LINE_SOFT, align="center"))
    ws.conditional_format(12, 5, 51, 5, {"type": "cell", "criteria": "<>", "value": 0,
                                         "format": k.f(bg_color=BLUSH, font_color=TERRA)})
    ws.conditional_format(12, 6, 51, 6, {"type": "text", "criteria": "containing",
                                         "value": "✓", "format": k.f(bg_color=MINT,
                                                                     font_color=GREEN, bold=True)})
    ws.conditional_format(12, 6, 51, 6, {"type": "text", "criteria": "containing",
                                         "value": "⚠", "format": k.f(bg_color="#FBE3D8",
                                                                     font_color=TERRA, bold=True)})
    ch = wb.add_chart({"type": "line"})
    ch.add_series({"name": "پز دفتر روزنامه", "categories": ["Pos", 12, 0, 41, 0],
                   "values": ["Pos", 12, 2, 41, 2], "line": {"color": GOLD, "width": 2.4},
                   "marker": {"type": "circle", "size": 4, "fill": {"color": WHITE},
                              "border": {"color": GOLD}}})
    ch.add_series({"name": "واریز پز به بانک", "categories": ["Pos", 12, 0, 41, 0],
                   "values": ["Pos", 12, 3, 41, 3], "line": {"color": FOREST, "width": 2.2},
                   "marker": {"type": "square", "size": 4, "fill": {"color": WHITE},
                              "border": {"color": FOREST}}})
    ch.set_title({"name": "پز دفتر در برابر واریز به بانک", "name_font": {"name": FA, "size": 11}})
    ch.set_legend({"position": "bottom", "font": {"name": FA, "size": 8}})
    ch.set_plotarea({"fill": {"color": "#FCFAF4"}})
    ch.set_size({"width": 620, "height": 240})
    ws.insert_chart("A55", ch, {"x_offset": 8, "y_offset": 6})
    k.note(ws, 53, last,
           "اختلاف بیش از ۱٬۰۰۰ ریال با ⚠ نمایش داده می‌شود. علت‌های رایج: واریزی پز در 02_Bank ثبت نشده، "
           "یا مبلغ پز در دفتر روزنامه اشتباه است. اعداد این نسخه نمونه‌اند.", bg=SAND, icon="warn")
    return ws


def sheet_banks(wb, k, banks):
    ws = wb.add_worksheet("Banks")
    last = 6
    k.setup(ws, last, 34, landscape=True, repeat=(11, 11), freeze=(13, 2))
    k.brand(ws, last, "وضعیت بانک‌ها",
            "مانده هر حساب از دفتر تراکنش‌های بانکی + سهم از کل", icon="bank")
    k.kpis(ws, 6, [
        ("جمع مانده", "=SUM($D$13:$D$22)", INT, FOREST, "money"),
        ("جمع دریافت", "=SUM($B$13:$B$22)", INT, GREEN, "transfer"),
        ("جمع پرداخت", "=SUM($C$13:$C$22)", INT, RED, "transfer"),
        ("خالص", "=$C$8-$E$8", INT_RED, BLUE, "chart"),
    ], last)
    heads = [("بانک", "Bank", 26), ("دریافت", "In", 16), ("پرداخت", "Out", 16),
             ("مانده", "Balance", 17), ("سهم", "Share", 9),
             ("مانده اولیه", "Opening", 15), ("حساب", "Acct", 10)]
    for i, (fa, en, w) in enumerate(heads):
        ws.set_column(i, i, w)
        ws.write(11, i, f"{fa}\n{en}",
                 k.f(font_name=FA, font_size=9, bold=True, font_color=WHITE,
                    bg_color=FOREST, border=1, border_color=FOREST, align="center",
                    text_wrap=True, valign="vcenter"))
    ws.set_row(11, 28)
    for i, (code, title, opening, accode, active, note) in enumerate(banks):
        r = 12 + i
        ws.write_string(r, 0, title, k.f(font_name=FA, font_size=9, bold=True,
                                         font_color=FOREST, bg_color=SAND_2, border=1,
                                         border_color=LINE, align="right"))
        ws.write_formula(r, 1, f'=SUMIF(F2!$D${F2_FIRST}:$D${F2_LAST},"{accode}",F2!$I${F2_FIRST}:$I${F2_LAST})',
                         k.f(font_name=EN, font_size=9, bg_color=MINT, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(r, 2, f'=SUMIF(F2!$D${F2_FIRST}:$D${F2_LAST},"{accode}",F2!$J${F2_FIRST}:$J${F2_LAST})',
                         k.f(font_name=EN, font_size=9, bg_color=BLUSH, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(r, 3, f'=$F{r+1}+$B{r+1}-$C{r+1}',
                         k.f(font_name=EN, font_size=10, bold=True, bg_color=CREAM,
                             border=1, border_color=LINE_SOFT, align="center",
                             num_format=INT_RED))
        ws.write_formula(r, 4, f'=IFERROR($D{r+1}/SUM($D$13:$D$22),0)',
                         k.f(font_name=EN, font_size=8.5, bg_color=WHITE, border=1,
                             border_color=LINE_SOFT, align="center", num_format=PCT))
        ws.write_number(r, 5, opening, k.f(font_name=EN, font_size=9, bg_color="#F6F5F0",
                                           border=1, border_color=LINE_SOFT,
                                           align="center", num_format=INT))
        ws.write_string(r, 6, accode, k.f(font_name=EN, font_size=8.5, bg_color="#F6F5F0",
                                          border=1, border_color=LINE_SOFT, align="center"))
    ch = wb.add_chart({"type": "bar"})
    ch.add_series({"name": "مانده", "categories": ["Banks", 12, 0, 21, 0],
                   "values": ["Banks", 12, 3, 21, 3], "fill": {"color": FOREST},
                   "data_labels": {"value": True, "num_format": "#,##0", "font": {"size": 7}}})
    ch.set_title({"name": "مانده هر بانک", "name_font": {"name": FA, "size": 11}})
    ch.set_legend({"none": True})
    ch.set_plotarea({"fill": {"color": "#FCFAF4"}})
    ch.set_size({"width": 480, "height": 250})
    ws.insert_chart("A26", ch, {"x_offset": 8, "y_offset": 6})
    return ws


def sheet_parties(wb, k, growers, buyers, bd):
    ws = wb.add_worksheet("Parties")
    last = 5
    k.setup(ws, last, 60, landscape=True, repeat=(11, 11), freeze=(13, 2))
    k.brand(ws, last, "طرف حساب‌ها",
            "باغداران، بارفروشان و باسکول‌های دیگران  ·  بدهکار، بستانکار و مانده", icon="people")
    heads = [("کد", "Code", 10), ("نام", "Name", 26), ("گروه", "Group", 14),
             ("بدهکار", "Debit", 16), ("بستانکار", "Credit", 16), ("مانده", "Balance", 16)]
    for i, (fa, en, w) in enumerate(heads):
        ws.set_column(i, i, w)
        ws.write(11, i, f"{fa}\n{en}",
                 k.f(font_name=FA, font_size=9, bold=True, font_color=WHITE,
                    bg_color=FOREST, border=1, border_color=FOREST, align="center",
                    text_wrap=True, valign="vcenter"))
    ws.set_row(11, 28)
    rows = [(c, n, "باغدار") for c, n in growers[:10]] + \
           [(c, n, "بارفروش") for c, n in buyers[:5]] + \
           [(c, n, "باسکول دیگران") for c, n in bd[:5]]
    for i, (code, name, grp) in enumerate(rows):
        r = 12 + i
        ws.write_string(r, 0, code, k.f(font_name=EN, font_size=8.5, bg_color=SAND_2,
                                        border=1, border_color=LINE, align="center"))
        ws.write_string(r, 1, name or "—", k.f(font_name=FA, font_size=9, bg_color=WHITE,
                                               border=1, border_color=LINE_SOFT,
                                               align="right"))
        ws.write_string(r, 2, grp, k.f(font_name=FA, font_size=8.5, bg_color=WHITE,
                                       border=1, border_color=LINE_SOFT, align="center"))
        ws.write_formula(r, 3, f'=SUMIF(F1!$D${F1_FIRST}:$D${F1_LAST},$A{r+1},F1!$L${F1_FIRST}:$L${F1_LAST})',
                         k.f(font_name=EN, font_size=9, bg_color=WHITE, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(r, 4, f'=SUMIF(F1!$D${F1_FIRST}:$D${F1_LAST},$A{r+1},F1!$M${F1_FIRST}:$M${F1_LAST})',
                         k.f(font_name=EN, font_size=9, bg_color=WHITE, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(r, 5, f'=$D{r+1}-$E{r+1}',
                         k.f(font_name=EN, font_size=9, bold=True, bg_color=CREAM,
                             border=1, border_color=LINE_SOFT, align="center",
                             num_format=INT_RED))
    ch = wb.add_chart({"type": "column"})
    ch.add_series({"name": "مانده", "categories": ["Parties", 12, 1, 31, 1],
                   "values": ["Parties", 12, 5, 31, 5], "fill": {"color": SAGE},
                   "data_labels": {"value": True, "num_format": "#,##0", "font": {"size": 6.5}}})
    ch.set_title({"name": "مانده طرف حساب‌ها", "name_font": {"name": FA, "size": 11}})
    ch.set_legend({"none": True})
    ch.set_plotarea({"fill": {"color": "#FCFAF4"}})
    ch.set_size({"width": 620, "height": 260})
    ws.insert_chart("A35", ch, {"x_offset": 8, "y_offset": 6})
    return ws


def sheet_dash(wb, k):
    ws = wb.add_worksheet("Dash")
    last = 7
    k.setup(ws, last, 62, landscape=True, repeat=None, freeze=None)
    k.brand(ws, last, "داشبورد مدیریتی مازندشت",
            "یک نگاه به وزن، مبلغ، بانک، پز و طرف حساب‌ها  ·  همه اعداد از ثبت روزانه و فایل بانک", icon="gauge")

    k.kpis(ws, 6, [
        ("کل وزن (kg)", "=Time!A8", INT, GREEN, "scale"),
        ("کل مبلغ", "=Time!C8", INT, GOLD, "money"),
        ("جمع بدهکار", "=Time!E8", INT, BLUE, "transfer"),
        ("اختلاف پز", "=Pos!G8", INT_RED, RED, "warn"),
    ], last)

    # ---- monthly chart data (helper block, hidden by width)
    k.section(ws, 8, last, "۰۱", "روند ماهانه — وزن و مبلغ", "از دفتر ثبت روزانه")
    ws.write(9, 0, "ماه", k.f(font_name=FA, font_size=8, bold=True, font_color=WHITE,
                              bg_color=FOREST, align="center"))
    ws.write(9, 1, "وزن (kg)", k.f(font_name=FA, font_size=8, bold=True, font_color=WHITE,
                                    bg_color=FOREST, align="center"))
    ws.write(9, 2, "مبلغ", k.f(font_name=FA, font_size=8, bold=True, font_color=WHITE,
                                bg_color=FOREST, align="center"))
    for m in range(12):
        r = 10 + m
        ws.write_formula(r, 0, f"=Time!$A${13+m}", k.f(font_name=EN, font_size=8, align="center"))
        ws.write_formula(r, 1, f"=Time!$C${13+m}", k.f(font_name=EN, font_size=8,
                                                        num_format=INT, align="center"))
        ws.write_formula(r, 2, f"=Time!$D${13+m}", k.f(font_name=EN, font_size=8,
                                                        num_format=INT, align="center"))
    ch = wb.add_chart({"type": "column"})
    ch.add_series({"name": "وزن (kg)", "categories": ["Dash", 9, 0, 20, 0],
                   "values": ["Dash", 9, 1, 20, 1], "fill": {"color": SAGE}, "gap": 50})
    ch.set_title({"name": "وزن ماهانه", "name_font": {"name": FA, "size": 11}})
    ln = wb.add_chart({"type": "line"})
    ln.add_series({"name": "مبلغ", "categories": ["Dash", 9, 0, 20, 0],
                   "values": ["Dash", 9, 2, 20, 2], "y2_axis": True,
                   "line": {"color": GOLD, "width": 2.2}})
    ch.combine(ln)
    ch.set_legend({"position": "bottom", "font": {"name": FA, "size": 8}})
    ch.set_plotarea({"fill": {"color": "#FCFAF4"}})
    ch.set_size({"width": 560, "height": 250})
    ws.insert_chart("D9", ch, {"x_offset": 4, "y_offset": 4})

    # ---- ops doughnut
    k.section(ws, 23, last, "۰۲", "ترکیب عملیات", "سهم هر نوع از مبلغ کل")
    for i, op in enumerate(["خرید مرکبات", "فروش بار (بارفروش)", "فروش متفرقه / صادراتی",
                            "آبگیری", "انعام مصطفی", "باسکول دیگران"]):
        r = 24 + i
        ws.write_formula(r, 0, f"=Ops!$A${13+i}", k.f(font_name=FA, font_size=8, align="right"))
        ws.write_formula(r, 1, f"=Ops!$D${13+i}", k.f(font_name=EN, font_size=8,
                                                       num_format=INT, align="center"))
    ch2 = wb.add_chart({"type": "doughnut"})
    ch2.add_series({"name": "ترکیب عملیات", "categories": ["Dash", 23, 0, 28, 0],
                    "values": ["Dash", 23, 1, 28, 1],
                    "points": [{"fill": {"color": c}} for c in
                               (FOREST, SAGE, GOLD, TERRA, AMBER, BLUE)],
                    "data_labels": {"percentage": True, "font": {"size": 8}}})
    ch2.set_title({"name": "سهم عملیات از مبلغ", "name_font": {"name": FA, "size": 11}})
    ch2.set_legend({"position": "right", "font": {"name": FA, "size": 8}})
    ch2.set_size({"width": 430, "height": 250})
    ws.insert_chart("D23", ch2, {"x_offset": 4, "y_offset": 4})

    # ---- bank positions mini table
    k.section(ws, 32, last, "۰۳", "مانده بانک‌ها", "از فایل تراکنش‌ها")
    for i, h in enumerate(["بانک", "دریافت", "پرداخت", "مانده"]):
        ws.write(33, i, h, k.f(font_name=FA, font_size=8.5, bold=True, font_color=WHITE,
                               bg_color=FOREST, align="center"))
    for i in range(5):
        r = 34 + i
        ws.write_formula(r, 0, f"=Banks!$A${13+i}", k.f(font_name=FA, font_size=8.5,
                                                         bg_color=SAND_2, border=1,
                                                         border_color=LINE, align="right"))
        ws.write_formula(r, 1, f"=Banks!$B${13+i}", k.f(font_name=EN, font_size=8.5,
                                                         num_format=INT, align="center"))
        ws.write_formula(r, 2, f"=Banks!$C${13+i}", k.f(font_name=EN, font_size=8.5,
                                                         num_format=INT, align="center"))
        ws.write_formula(r, 3, f"=Banks!$D${13+i}", k.f(font_name=EN, font_size=8.5, bold=True,
                                                         num_format=INT, bg_color=CREAM,
                                                         align="center"))

    # ---- alerts + nav
    k.section(ws, 40, last, "۰۴", "هشدارها", "مواردی که باید همین امروز بررسی شود")
    alerts = [
        ("ردیف‌های ناقص باسکول", '=SUM(F1!$L$10:$L$1808)-SUM(F1!$M$10:$M$1808)', "اختلاف بدهکار/بستانکار دفتر"),
        ("اختلاف پز", "=Pos!G8", "پز دفتر روزنامه منهای واریز پز به بانک"),
        ("مانده منفی بانک", '=COUNTIF(Banks!$D$13:$D$22,"<0")', "تعداد حساب‌هایی که مانده منفی دارند"),
    ]
    for i, (lab, formula, hint) in enumerate(alerts):
        r = 41 + i
        ws.merge_range(r, 0, r, 2, lab, k.f(font_name=FA, font_size=9.5, bold=True,
                                            bg_color=SAND_2, border=1, border_color=LINE,
                                            align="right", valign="vcenter"))
        ws.merge_range(r, 3, r, 4, "", k.f(font_name=EN, font_size=12, bold=True,
                                           font_color=TERRA, bg_color=CREAM, border=1,
                                           border_color=LINE, align="center",
                                           num_format=INT))
        ws.write_formula(r, 3, formula, k.f(font_name=EN, font_size=12, bold=True,
                                            font_color=TERRA, bg_color=CREAM, border=1,
                                            border_color=LINE, align="center", num_format=INT))
        ws.merge_range(r, 5, r, last, hint, k.f(font_name=FA, font_size=8, font_color=MUTED,
                                                align="right", valign="vcenter"))
        ws.set_row(r, 22)

    k.section(ws, 45, last, "۰۵", "شیت‌های این داشبورد", "روی هر کارت کلیک کنید")
    k.nav(ws, 46, [
        ("پز و واریزی  Pos", "کنترل روزانه پز دفتر روزنامه در برابر واریز پز به بانک", "internal:Pos!A1", "card", GOLD),
        ("بانک‌ها  Banks", "مانده و سهم هر بانک از گردش", "internal:Banks!A1", "bank", FOREST),
        ("روند  Time", "وزن و مبلغ ماه به ماه", "internal:Time!A1", "chart", SAGE),
        ("عملیات  Ops", "تحلیل نوع عملیات", "internal:Ops!A1", "gauge", BLUE),
        ("طرف حساب‌ها  Parties", "باغدار، بارفروش و باسکول دیگران", "internal:Parties!A1", "people", TERRA),
        ("راهنما  Help", "راهنمای کامل داشبورد و اتصال‌ها", "internal:Help!A1", "help", "#8E6C88"),
    ], last, per_row=3)

    k.section(ws, 52, last, "۰۶", "فایل‌های همراه",
              "همه در همین پوشه — داشبورد از دو فایل اول داده می‌خواند")
    k.nav(ws, 53, [
        ("دفتر ثبت  01_Entry", "منبع اصلی داشبورد: شیت Book همین فایل است",
         "external:01_Entry.xlsx", "list", FOREST),
        ("بانک  02_Bank", "منبع شیت Banks و پز؛ شیت Post",
         "external:02_Bank.xlsx", "bank", GOLD),
        ("فایل مادر  00_Master", "حساب‌ها، بانک‌ها و فهرست‌های کشویی",
         "external:00_Master.xlsx", "list", SAGE),
        ("تراز  05_Bal", "تراز آزمایشی و گردش حساب‌ها",
         "external:05_Bal.xlsx", "scale", BLUE),
        ("فاکتورها  04_Inv", "دفتر فاکتورهای فروش و خرید",
         "external:04_Inv.xlsx", "cheque", TERRA),
        ("گزارش‌ها  03_Report", "سود و زیان، گردش حساب و عملکرد",
         "external:03_Report.xlsx", "gauge", "#8E6C88"),
    ], last, per_row=3)
    return ws


def sheet_help(wb, k):
    ws = wb.add_worksheet("Help")
    last = 7
    k.setup(ws, last, 70, landscape=False, repeat=(7, 7))
    k.brand(ws, last, "راهنمای داشبورد",
            "این فایل چه می‌خواند، چگونه به‌روز می‌شود و هر نمودار چه می‌گوید", icon="help")
    blocks = [
        ("۰۱", "داشبورد از کجا می‌خواند؟", [
            ("p", "داشبورد به دو فایل کنار خودش وصل است: دفتر ثبت روزانه (01_Entry → شیت Book) و "
                  "دفتر تراکنش‌های بانکی (02_Bank → شیت Post). دو شیت مخفی F1 و F2 همین دو منبع را آینه می‌کنند."),
            ("b", "برای دیدن اعداد تازه: هر دو فایل را ذخیره کنید، بعد داشبورد را باز کنید و Update Links بزنید."),
            ("b", "اگر فایلی را جابه‌جا کردید، همه فایل‌ها را با هم جابه‌جا کنید (لینک‌ها نسبی‌اند)."),
        ]),
        ("۰۲", "کارت‌های بالای صفحه", [
            ("b", "کل وزن (kg): جمع وزن خالص همه ردیف‌های باسکول در کل دوره — همان عدد شیت Time."),
            ("b", "کل مبلغ: جمع مبلغ تمام ردیف‌های دفتر (خرید، فروش، هزینه و …)."),
            ("b", "جمع بدهکار: جمع ستون بدهکار دفتر؛ در ثبت دوطرفه با جمع بستانکار برابر است."),
            ("b", "اختلاف پز: پز ثبت‌شده در دفتر منهای واریز پز به بانک؛ باید نزدیک صفر باشد."),
        ]),
        ("۰۳", "شیت Pos", [
            ("p", "هر ردیف یک تاریخ است. اگر روز واریز با روز پز یکی نیست، تاریخ واریز را در ستون B بنویسید."),
            ("b", "ستون «سایر واریزها» واریزهایی است که از راه واسط X011 به بانک رسیده و مربوط به پز نیست."),
            ("b", "برای ماه جدید، یک تاریخ جدید در ستون A اضافه کنید."),
        ]),
        ("۰۴", "شیت Time و Ops", [
            ("b", "Time روند ماهانه وزن و مبلغ است؛ ماه از تاریخ شمسی استخراج می‌شود."),
            ("b", "Ops نشان می‌دهد هر نوع عملیات چقدر وزن و مبلغ جابه‌جا کرده و سهمش چقدر است."),
        ]),
        ("۰۵", "شیت Banks و Parties", [
            ("b", "Banks مانده هر بانک را از دفتر تراکنش‌ها می‌سازد: مانده اولیه + دریافت − پرداخت."),
            ("b", "Parties مانده باغداران، بارفروشان و باسکول‌های دیگران را نشان می‌دهد."),
        ]),
        ("۰۶", "چاپ", [
            ("p", "همه شیت‌ها A4، راست‌به‌چپ، با سربرگ/فوتر و در عرض یک صفحه تنظیم شده‌اند."),
        ]),
    ]
    r = 6
    for num, title, items in blocks:
        k.section(ws, r, last, num, title)
        r += 1
        for kind, text in items:
            if kind == "p":
                ws.merge_range(r, 0, r, last, text,
                               k.f(font_name=FA, font_size=10, font_color=INK,
                                   align="right", valign="top", text_wrap=True))
                ws.set_row(r, 32)
                r += 1
            else:
                ws.merge_range(r, 0, r, last, "     •  " + text,
                               k.f(font_name=FA, font_size=10, font_color=FOREST,
                                   bg_color=CREAM, align="right", valign="vcenter",
                                   text_wrap=True, border=1, border_color=LINE_SOFT))
                ws.set_row(r, 26)
                r += 1
        r += 1
    ws.print_area(0, 0, r, last)
    return ws


def sheet_sync(wb, k, banks, growers, buyers, bd, groups):
    ws = wb.add_worksheet("Sync")
    k.setup(ws, 8, 60, landscape=True, repeat=(8, 8), freeze=(9, 2))
    ws.merge_range(0, 0, 0, 8, "مرجع قفل‌شده — کپی از فایل مادر و طرف حساب‌های کلیدی",
                   k.f(font_name=FA_TITLE, font_size=14, bold=True, font_color=FOREST,
                       bg_color=SAND_2, align="right"))
    ws.merge_range(1, 0, 1, 8, "این شیت فقط برای نمایش است؛ تغییر در آن مجاز نیست.",
                   k.f(font_name=FA, font_size=9, font_color=MUTED, align="right"))
    ws.write(5, 0, "بانک", k.f(font_name=FA, font_size=9, bold=True, font_color=WHITE,
                                bg_color=FOREST, align="center"))
    ws.write(5, 1, "حساب", k.f(font_name=FA, font_size=9, bold=True, font_color=WHITE,
                                bg_color=FOREST, align="center"))
    ws.write(5, 3, "باغدار", k.f(font_name=FA, font_size=9, bold=True, font_color=WHITE,
                                  bg_color=FOREST, align="center"))
    ws.write(5, 4, "بارفروش", k.f(font_name=FA, font_size=9, bold=True, font_color=WHITE,
                                   bg_color=FOREST, align="center"))
    ws.write(5, 5, "باسکول دیگران", k.f(font_name=FA, font_size=9, bold=True, font_color=WHITE,
                                         bg_color=FOREST, align="center"))
    for i, b in enumerate(banks):
        ws.write(6 + i, 0, b[1], k.f(font_name=FA, font_size=9, align="right"))
        ws.write(6 + i, 1, b[3], k.f(font_name=EN, font_size=9, align="center"))
    for i, (c, n) in enumerate(growers[:10]):
        ws.write(6 + i, 3, f"{c} – {n}", k.f(font_name=FA, font_size=9, align="right"))
    for i, (c, n) in enumerate(buyers[:5]):
        ws.write(6 + i, 4, f"{c} – {n}", k.f(font_name=FA, font_size=9, align="right"))
    for i, (c, n) in enumerate(bd[:5]):
        ws.write(6 + i, 5, f"{c} – {n or '—'}", k.f(font_name=FA, font_size=9, align="right"))
    ws.protect("", {"select_locked_cells": True})
    return ws


def build(master_path, bank_path, entry_path, out_path, feed_values=None):
    acc, banks, growers, buyers, bd, groups = read_master(master_path)
    vals = feed_values or {}
    wb = xlsxwriter.Workbook(out_path, {"default_format_properties": {
        "font_name": FA, "font_size": 10}})
    wb.set_properties({"title": "مازندشت — داشبورد مدیریتی",
                       "subject": "داشبورد و تحلیل", "author": "MAZANDASHT",
                       "company": "سورتینگ و بسته‌بندی مرکبات مازندشت"})
    k = Kit(wb)
    sheet_dash(wb, k)
    sheet_pos(wb, k, vals.get("dates") or [])
    sheet_banks(wb, k, banks)
    sheet_time(wb, k)
    sheet_ops(wb, k, vals.get("ops") or [])
    sheet_parties(wb, k, growers, buyers, bd)
    sheet_prod(wb, k, vals.get("F1") or {})
    sheet_f1(wb, k, vals.get("F1") or {})
    sheet_f2(wb, k, vals.get("F2") or {})
    sheet_sync(wb, k, banks, growers, buyers, bd, groups)
    sheet_help(wb, k)
    wb.worksheets()[0].activate()
    wb.close()
    return out_path


def read_master(path):
    wb = openpyxl.load_workbook(path)
    acc = wb["Acc"]
    rows = []
    for r in range(10, 260):
        code = acc.cell(r, 1).value
        if code:
            rows.append((str(code), acc.cell(r, 4).value, acc.cell(r, 2).value))
    banks = []
    bnk = wb["Bnk"]
    for r in range(10, 20):
        if bnk.cell(r, 1).value:
            banks.append((bnk.cell(r, 1).value, bnk.cell(r, 2).value, bnk.cell(r, 3).value or 0,
                          bnk.cell(r, 4).value, bnk.cell(r, 6).value or "بله",
                          bnk.cell(r, 7).value or ""))
    ops = []
    op = wb["Op"]
    for r in range(10, 34):
        if op.cell(r, 2).value:
            ops.append(op.cell(r, 2).value)
    groups = []
    grp = wb["Grp"]
    for r in range(10, 55):
        g = grp.cell(r, 1).value
        if g and g not in groups:
            groups.append(g)
    wb.close()
    growers = [(c, n) for c, n, g in rows if g == "باغدار"][:10]
    buyers = [(c, n) for c, n, g in rows if g == "بارفروش"][:5]
    bd = [(c, n) for c, n, g in rows if g == "باسکول دیگران"][:5]
    return rows, banks, growers, buyers, bd, groups
