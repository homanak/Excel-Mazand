# -*- coding: utf-8 -*-
"""01_Entry — capacity, dropdown hygiene, tooltips and three new print sheets.

* grows Scale / Crate / Pay / Book to their final size (see ``grow.py``)
* re-points every data validation to the new last row
* adds ``Inv``    — دفتر فاکتورها (خرید، فروش، باسکول، سبد) with a print tick
* adds ``Print``  — فاکتور چاپی A4: only the ticked rows, with totals + signatures
* adds ``GRep``   — گزارش چاپی گروه‌های حساب + روند ماهانه
* adds crisp Persian tooltips on every input column of the three entry sheets
"""
import openpyxl
from openpyxl.comments import Comment
from openpyxl.worksheet.datavalidation import DataValidation

import grow
import look as L
from look import box, fill

INV_LAST = 2009          # Inv!10 .. Inv!2009 → 2,000 invoices
PRINT_ROWS = 14          # lines on one printed invoice


# ------------------------------------------------------------------ helpers
def _sheet(wb, title, index=None):
    ws = wb.create_sheet(title) if index is None else wb.create_sheet(title, index)
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False
    return ws


def dv_resize(ws, mapping):
    """Rewrite ``<col>10:<col>159`` ranges of the validations of one sheet."""
    import re
    out = []
    for dv in ws.data_validations.dataValidation:
        sq = str(dv.sqref)
        m = re.fullmatch(r"([A-Z]{1,3})10:\1(\d+)", sq)
        if m and int(m.group(2)) in mapping:
            dv.sqref = f"{m.group(1)}10:{m.group(1)}{mapping[int(m.group(2))]}"
        out.append(dv)
    return len(out)


def tooltips(ws, tips, row=8):
    for ref, text in tips.items():
        cell = ws[f"{ref}{row}"]
        cell.comment = Comment(text, "MAZANDASHT", width=300, height=130)


# ------------------------------------------------------------------- Invoice
def add_inv(wb, style_src="Scale"):
    ws = _sheet(wb, "Inv", 4)          # after Pay
    last = 20
    L.brand(ws, last,
            "دفتر فاکتورها — خرید، فروش، باسکول و سبد",
            "هر ردیف یک قلم فاکتور  ·  با تیک ستون «چاپ» در شیت Print چاپ می‌شود  ·  "
            "کد طرف حساب را تایپ کنید؛ نام و گروه خودکار می‌آید")
    L.kpis(ws, 6, last, [
        ("تعداد فاکتور", '=COUNTA($B$10:$B$2009)', L.INT, L.FOREST),
        ("جمع مبلغ", '=SUM($O$10:$O$2009)', L.INT, L.GOLD),
        ("ردیف‌های انتخاب‌شده برای چاپ", '=COUNTIF($S$10:$S$2009,"✔")', L.INT, L.GREEN),
        ("مبلغ انتخاب‌شده", '=SUMIF($S$10:$S$2009,"✔",$O$10:$O$2009)', L.INT, L.TERRA),
    ])

    L.section(ws, 8, last, "۰۱", "فاکتورها", "— تاریخ، نوع، طرف حساب و اقلام")
    cols = [
        ("ردیف", "No.", 6, "خودکار"),
        ("شماره فاکتور", "Invoice", 13, "شماره فاکتور کاغذی؛ یکتا باشد"),
        ("تاریخ", "Date", 11, "قالب 1405/07/10 — ارقام لاتین"),
        ("نوع عملیات", "Operation", 17, "از فهرست کشویی انتخاب کنید (خرید مرکبات، فروش بار، …)"),
        ("کد طرف حساب", "Party", 11, "کد را تایپ کنید: C001… باغدار، B001… بارفروش، S001… فروشنده سبد"),
        ("نام حساب", "Name", 20, "خودکار از فایل مادر"),
        ("گروه", "Group", 11, "خودکار"),
        ("کالا / شرح", "Item", 16, "محصول از فهرست؛ برای سایر موارد تایپ کنید"),
        ("سبد", "Crate", 12, "نوع سبد از فهرست"),
        ("تعداد سبد", "Qty", 9, "تعداد سبد"),
        ("وزن (kg)", "Weight", 10, "اگر فاکتور وزنی است"),
        ("فی", "Price", 11, "قیمت واحد (ریال)"),
        ("مبلغ", "Amount", 14, "وزن×فی (یا تعداد×فی) — خودکار"),
        ("تخفیف", "Disc.", 10, "اختیاری"),
        ("مبلغ خالص", "Net", 14, "خودکار"),
        ("روش پرداخت", "Pay", 11, "نقد، چک، انتقال بانکی، کارت/پوز"),
        ("سررسید", "Due", 11, "برای چک و نسیه لازم است"),
        ("سند مرتبط", "Doc", 12, "شماره سند دفتر روزنامه (BS-001، PV-003 …)"),
        ("چاپ", "Print", 7, "✔ بزنید تا این ردیف در فاکتور چاپی بیاید"),
        ("#", "", 4, "سریال چاپ — دست نزنید"),
    ]
    L.thead(ws, 9, cols)

    for i in range(INV_LAST - 9):
        r = 10 + i
        ws.row_dimensions[r].height = 19
        L.put(ws, f"A{r}", f'=IF($B{r}="","",COUNTA($B$10:$B{r}))', name=L.EN, size=8.5,
              bg=L.GREY, color=L.MUTED)
        for col in ("B", "C", "D", "E", "P", "Q", "R", "S"):
            L.put(ws, f"{col}{r}", None, name=L.EN if col in "BCQ" else L.FA, size=8.5,
                  bg=L.INPUT, h="center" if col != "E" else "right")
        L.put(ws, f"F{r}", f'=IF($E{r}="","",IFERROR(VLOOKUP($E{r},AcctTbl,2,0),"کد ناشناخته"))',
              name=L.FA, size=8.5, bg=L.SAND2, h="right")
        L.put(ws, f"G{r}", f'=IF($E{r}="","",IFERROR(VLOOKUP($E{r},AcctTbl,3,0),""))',
              name=L.FA, size=8.5, bg=L.SAND2)
        for col in ("H", "I"):
            L.put(ws, f"{col}{r}", None, name=L.FA, size=8.5, bg=L.INPUT, h="right")
        for col in ("J", "K", "L", "N"):
            L.put(ws, f"{col}{r}", None, name=L.EN, size=8.5, bg=L.INPUT, fmt=L.INT)
        L.put(ws, f"M{r}",
              f'=IF($B{r}="","",ROUND(IF(N($K{r})=0,N($J{r}),$K{r})*N($L{r}),0))',
              name=L.EN, size=8.5, bg=L.CREAM, fmt=L.INT, bold=True)
        L.put(ws, f"O{r}", f'=IF($B{r}="","",N($M{r})-N($N{r}))', name=L.EN, size=8.5,
              bg=L.GOLD_L, fmt=L.INT, bold=True)
        L.put(ws, f"T{r}", f'=IF($S{r}="✔",N(MAX($T$9:T{r-1}))+1,"")', name=L.EN,
              size=8, bg=L.GREY, color=L.MUTED)

    # ---- یک فاکتور نمونه، تیک‌خورده، تا شیت Print از همان روز اول کار کند
    demo = [
        # (row, invoice, date, op, party, item, basket, qty, weight, price, method, due, doc, tick)
        (10, "INV-1405-001", "1405/07/12", "فروش بار (بارفروش)", "SL001",
         "پرتقال تامسون", "BK03 – ۹ک مشکی", 0, 2500, 58000,
         "انتقال بانکی", "1405/08/12", "BS-006", "✔"),
        (11, "INV-1405-001", "1405/07/12", "فروش بار (بارفروش)", "SL001",
         "پرتقال تامسون", "BK03 – ۹ک مشکی", 0, 1600, 60000,
         "انتقال بانکی", "1405/08/12", "BS-015", "✔"),
        (12, "INV-1405-002", "1405/07/13", "خرید مرکبات", "C002",
         "پرتقال خونی", "BK01 – ۶ک مشکی", 0, 800, 35000,
         "چک", "1405/08/20", "BS-004", None),
    ]
    if INV_LAST >= 40:
        for (r, inv_no, day, op, party, item, basket, qty, kg, price,
             method, due, doc, tick) in demo:
            for col, val in (("B", inv_no), ("C", day), ("D", op), ("E", party),
                             ("H", item), ("I", basket), ("J", qty), ("K", kg),
                             ("L", price), ("P", method), ("Q", due), ("R", doc),
                             ("S", tick)):
                ws[f"{col}{r}"] = val

    # totals
    tot = INV_LAST + 1
    ws.merge_cells(f"A{tot}:I{tot}")
    L.put(ws, f"A{tot}", "جمع کل فاکتورها", name=L.FA, size=10, bold=True, color=L.WHITE,
          bg=L.GOLD)
    for col in ("J", "K", "L", "M", "N", "O"):
        L.put(ws, f"{col}{tot}", f"=SUM({col}10:{col}{INV_LAST})", name=L.EN, size=10,
              bold=True, color=L.WHITE, bg=L.GOLD, fmt=L.INT)
    L.put(ws, f"P{tot}", f'=COUNTIF($S$10:$S$2009,"✔")&" چاپ"', name=L.FA, size=9.5,
          bold=True, color=L.WHITE, bg=L.GOLD)

    # validations
    def dv(formula, rng, title="", msg="", style="stop"):
        d = DataValidation(type="list", formula1=formula, allow_blank=True, showErrorMessage=True,
                           errorTitle=title or "مقدار معتبر نیست",
                           error=msg or "از فهرست کشویی انتخاب کنید.")
        ws.add_data_validation(d)
        d.add(rng)

    dv("ListBaskolOps", f"D10:D{INV_LAST}", "نوع عملیات", "از فهرست عملیات باسکول انتخاب کنید.")
    dv("ListAllCodes", f"E10:E{INV_LAST}", "کد طرف حساب", "کد را از فهرست انتخاب یا تایپ کنید (C001…).")
    dv("ListProducts", f"H10:H{INV_LAST}", "کالا", "محصول را انتخاب کنید یا شرح دلخواه بنویسید.")
    dv("ListBaskets", f"I10:I{INV_LAST}", "سبد", "نوع سبد را انتخاب کنید.")
    dv("ListPayMethods", f"P10:P{INV_LAST}", "روش پرداخت", "نقد، چک، انتقال بانکی، کارت/پوز …")
    dv('"✔"', f"S10:S{INV_LAST}", "چاپ", "برای چاپ، ✔ بگذارید؛ برای حذف، خانه را خالی کنید.")

    ws.column_dimensions["T"].hidden = True
    ws.freeze_panes = "F10"
    ws.auto_filter.ref = f"A9:T{INV_LAST}"
    L.page(ws, last, tot, landscape=True, header_rows="9:9")
    ws.print_area = f"A1:T{tot + 2}"
    L.note(ws, tot + 2, last,
           "کد طرف حساب را تایپ کنید (کشو باز می‌شود ولی تایپ هم آزاد است)؛ اگر کد در فایل مادر نباشد "
           "«کد ناشناخته» می‌شود. ردیف‌های تیک‌خورده در ستون سیاه‌رنگ «چاپ» به ترتیب در شیت Print چاپ می‌شوند.",
           bg=L.MINT, h=32)
    return ws


# -------------------------------------------------------------------- Print
def add_print(wb):
    ws = _sheet(wb, "Print", 5)
    last = 8
    for col, w in zip("ABCDEFGH", (5, 30, 9, 12, 13, 16, 14, 3)):
        ws.column_dimensions[col].width = w
    L.brand(ws, last, "فاکتور چاپی",
            "فقط ردیف‌هایی که در شیت Inv ستون «چاپ» را ✔ زده‌اید  ·  آماده چاپ A4")

    # --- head card
    ws.merge_cells("A6:E6")
    L.put(ws, "A6", "مشخصات فاکتور", name=L.FA, size=11, bold=True, color=L.WHITE, bg=L.FOREST)
    ws.merge_cells("F6:H6")
    L.put(ws, "F6", "خریدار / فروشنده", name=L.FA, size=11, bold=True, color=L.WHITE,
          bg=L.FOREST)

    cards = [
        (7, "A", "شماره فاکتور", "B", "D",
         '=IFERROR(INDEX(Inv!$B$10:$B$2009,MATCH("✔",Inv!$S$10:$S$2009,0)),"")'),
        (7, "E", "تاریخ", "F", "H",
         '=IFERROR(INDEX(Inv!$C$10:$C$2009,MATCH("✔",Inv!$S$10:$S$2009,0)),"")'),
        (8, "A", "طرف حساب", "B", "D",
         '=IFERROR(INDEX(Inv!$E$10:$E$2009,MATCH("✔",Inv!$S$10:$S$2009,0))&" — "'
         '&INDEX(Inv!$F$10:$F$2009,MATCH("✔",Inv!$S$10:$S$2009,0)),"")'),
        (8, "E", "نوع / پرداخت", "F", "H",
         '=IFERROR(INDEX(Inv!$D$10:$D$2009,MATCH("✔",Inv!$S$10:$S$2009,0))&" / "'
         '&INDEX(Inv!$P$10:$P$2009,MATCH("✔",Inv!$S$10:$S$2009,0)),"")'),
    ]
    for row, lc, label, v0, v1, formula in cards:
        L.put(ws, f"{lc}{row}", label, name=L.FA, size=8.5, bold=True, color=L.MUTED,
              bg=L.SAND2, h="right")
        ws.merge_cells(f"{v0}{row}:{v1}{row}")
        L.put(ws, f"{v0}{row}", formula, name=L.FA, size=10, bold=True, color=L.FOREST,
              bg=L.CREAM)
        for col in range(ws[f"{v0}{row}"].column, ws[f"{v1}{row}"].column + 1):
            ws.cell(row, col).fill = fill(L.CREAM)
            ws.cell(row, col).border = box(L.LINE)

    L.section(ws, 10, last, "۰۱", "اقلام فاکتور", "")

    cols = [("ردیف", "No.", 5, None), ("شرح کالا / خدمت", "Description", 30, None),
            ("تعداد سبد", "Qty", 9, None), ("وزن (kg)", "Weight", 12, None),
            ("فی", "Price", 13, None), ("مبلغ", "Amount", 16, None),
            ("گروه", "Group", 14, None), ("", "", 3, None)]
    L.thead(ws, 11, cols)

    for i in range(PRINT_ROWS):
        r = 12 + i
        ws.row_dimensions[r].height = 22
        L.put(ws, f"A{r}", i + 1, name=L.EN, size=8.5, bg=L.SAND2)
        L.put(ws, f"B{r}",
              f'=IFERROR(INDEX(Inv!$H$10:$H$2009,MATCH({i+1},Inv!$T$10:$T$2009,0)),"")',
              name=L.FA, size=9, bg=L.WHITE, h="right")
        L.put(ws, f"C{r}",
              f'=IFERROR(IF(INDEX(Inv!$J$10:$J$2009,MATCH({i+1},Inv!$T$10:$T$2009,0))=0,"",'
              f'INDEX(Inv!$J$10:$J$2009,MATCH({i+1},Inv!$T$10:$T$2009,0))),"")',
              name=L.EN, size=9, bg=L.WHITE, fmt=L.INT)
        L.put(ws, f"D{r}",
              f'=IFERROR(INDEX(Inv!$K$10:$K$2009,MATCH({i+1},Inv!$T$10:$T$2009,0)),"")',
              name=L.EN, size=9, bg=L.WHITE, fmt=L.DEC)
        L.put(ws, f"E{r}",
              f'=IFERROR(INDEX(Inv!$L$10:$L$2009,MATCH({i+1},Inv!$T$10:$T$2009,0)),"")',
              name=L.EN, size=9, bg=L.WHITE, fmt=L.INT)
        L.put(ws, f"F{r}",
              f'=IFERROR(INDEX(Inv!$O$10:$O$2009,MATCH({i+1},Inv!$T$10:$T$2009,0)),"")',
              name=L.EN, size=9.5, bold=True, bg=L.CREAM, fmt=L.INT)
        L.put(ws, f"G{r}",
              f'=IFERROR(INDEX(Inv!$G$10:$G$2009,MATCH({i+1},Inv!$T$10:$T$2009,0)),"")',
              name=L.FA, size=8.5, bg=L.WHITE)

    tot = 12 + PRINT_ROWS
    ws.merge_cells(f"A{tot}:E{tot}")
    L.put(ws, f"A{tot}", "جمع کل فاکتور", name=L.FA, size=11, bold=True, color=L.WHITE, bg=L.GOLD,
          h="right")
    L.put(ws, f"F{tot}", f"=SUM(F12:F{tot-1})", name=L.EN, size=12, bold=True, color=L.WHITE,
          bg=L.GOLD, fmt=L.INT)
    L.put(ws, f"G{tot}", f'=COUNTIF(Inv!$S$10:$S$2009,"✔")&" ردیف"', name=L.FA, size=9,
          bold=True, color=L.WHITE, bg=L.GOLD)

    # signatures
    s = tot + 2
    for i, (c0, c1, label) in enumerate([("A", "C", "مهر و امضای فروشنده"),
                                         ("D", "F", "امضای خریدار"),
                                         ("G", "H", "تاریخ و تأیید")]):
        ws.merge_cells(f"{c0}{s}:{c1}{s}")
        L.put(ws, f"{c0}{s}", label, name=L.FA, size=9, bold=True, color=L.MUTED, bg=L.SAND2)
        ws.merge_cells(f"{c0}{s+1}:{c1}{s+1}")
        L.put(ws, f"{c0}{s+1}", "", bg=L.WHITE)
        ws.row_dimensions[s + 1].height = 54

    L.note(ws, s + 3, last,
           "برای چاپ: در شیت Inv ردیف‌های مورد نظر را در ستون «چاپ» با ✔ علامت بزنید؛ همین صفحه فقط "
           "همان ردیف‌ها را به ترتیب نشان می‌دهد. برای فاکتور بعدی، تیک‌های قبلی را بردارید.",
           bg=L.MINT, h=30)
    L.page(ws, last, s + 3, landscape=False)
    ws.print_area = f"A1:H{s+3}"
    return ws


# --------------------------------------------------------------------- GRep
def add_grep(wb, groups=None, book_last=None):
    groups = groups or ["باغدار", "بارفروش", "فروشنده سبد", "تأمین‌کننده", "کارمند",
                        "سایر حساب", "باسکول دیگران", "دارایی ثابت", "وام و تسهیلات",
                        "سرمایه و آورده صاحبان", "بانک و صندوق"]
    book_last = book_last or grow.BOOK_LAST
    ws = _sheet(wb, "GRep", 6)
    last = 7
    L.brand(ws, last, "گزارش گروه‌های حساب",
            "گردش و مانده هر گروه در دفتر روزنامه  ·  آماده چاپ A4  ·  برای گزارش یک گروه، "
            "ردیف همان گروه را جدا کنید")
    L.kpis(ws, 6, last, [
        ("جمع بدهکار", '=SUM($B$13:$B$23)', L.INT, L.FOREST),
        ("جمع بستانکار", '=SUM($C$13:$C$23)', L.INT, L.GOLD),
        ("مانده", '=$B$7-$C$7', L.INT_RED, L.TERRA),
    ])
    L.section(ws, 11, last, "۰۱", "گردش گروه‌ها", "— بدهکار، بستانکار، مانده و سهم")
    cols = [("گروه حساب", "Group", 26, None), ("بدهکار", "Debit", 17, None),
            ("بستانکار", "Credit", 17, None), ("مانده", "Balance", 17, None),
            ("تعداد سطر", "Lines", 11, None), ("سهم", "Share", 10, None),
            ("وضعیت", "Status", 13, None)]
    L.thead(ws, 12, cols)
    for i, g in enumerate(groups):
        r = 13 + i
        ws.row_dimensions[r].height = 22
        L.put(ws, f"A{r}", g, name=L.FA, size=9.5, bold=True, bg=L.SAND2, h="right")
        L.put(ws, f"B{r}", f'=SUMIF(Book!$O$10:$O${book_last},$A{r},Book!$L$10:$L${book_last})',
              name=L.EN, size=9, bg=L.WHITE, fmt=L.INT)
        L.put(ws, f"C{r}", f'=SUMIF(Book!$O$10:$O${book_last},$A{r},Book!$M$10:$M${book_last})',
              name=L.EN, size=9, bg=L.WHITE, fmt=L.INT)
        L.put(ws, f"D{r}", f"=$B{r}-$C{r}", name=L.EN, size=9.5, bold=True, bg=L.CREAM,
              fmt=L.INT_RED)
        L.put(ws, f"E{r}", f'=COUNTIF(Book!$O$10:$O${book_last},$A{r})', name=L.EN, size=9,
              bg=L.WHITE, fmt=L.INT)
        L.put(ws, f"F{r}", f'=IFERROR($D{r}/SUM($D$13:$D$23),0)', name=L.EN, size=8.5,
              bg=L.WHITE, fmt=L.PCT)
        L.put(ws, f"G{r}", f'=IF(ROUND($D{r},0)=0,"—",IF($D{r}>0,"بدهکار","بستانکار"))',
              name=L.FA, size=8.5, bg=L.WHITE)
    tot = 13 + len(groups)
    ws.merge_cells(f"A{tot}:A{tot}")
    L.put(ws, f"A{tot}", "جمع", name=L.FA, size=10, bold=True, color=L.WHITE, bg=L.GOLD)
    for col in "BCDE":
        L.put(ws, f"{col}{tot}", f"=SUM({col}13:{col}{tot-1})", name=L.EN, size=10, bold=True,
              color=L.WHITE, bg=L.GOLD, fmt=L.INT)
    L.put(ws, f"F{tot}", f"=SUM(F13:F{tot-1})", name=L.EN, size=10, bold=True, color=L.WHITE,
          bg=L.GOLD, fmt=L.PCT)
    L.put(ws, f"G{tot}", f'=IF(ROUND($B{tot}-$C{tot},0)=0,"✓ تراز","⚠")', name=L.FA,
          size=9.5, bold=True, color=L.WHITE, bg=L.GOLD)

    # share bars — a printable picture of where the money sits
    ws.column_dimensions["H"].hidden = True
    ws.column_dimensions["H"].width = 8
    for i in range(len(groups)):
        r = 13 + i
        L.put(ws, f"H{r}", f"=ABS($D{r})", name=L.EN, size=8, color=L.GREY, bg=L.GREY)

    m0 = tot + 2
    L.section(ws, m0, last, "۰۲", "نمای سهم گروه‌ها", "— طول میله با اندازه مانده هر گروه متناسب است")
    head = ws.cell(m0 + 1, 1, "گروه حساب")
    L.st(head, name=L.FA, size=9, bold=True, color=L.WHITE, bg=L.FOREST, border=L.FOREST)
    head2 = ws.cell(m0 + 1, 2, "مانده")
    L.st(head2, name=L.FA, size=9, bold=True, color=L.WHITE, bg=L.FOREST, border=L.FOREST)
    ws.merge_cells(start_row=m0 + 1, start_column=3, end_row=m0 + 1, end_column=7)
    head3 = ws.cell(m0 + 1, 3, "نمای نسبی")
    L.st(head3, name=L.FA, size=9, bold=True, color=L.WHITE, bg=L.FOREST, border=L.FOREST)
    for c in range(3, 8):
        ws.cell(m0 + 1, c).fill = fill(L.FOREST)
    for i, g in enumerate(groups):
        r = m0 + 2 + i
        L.put(ws, f"A{r}", f"=$A${13 + i}", name=L.FA, size=9, bg=L.WHITE, h="right")
        L.put(ws, f"B{r}", f"=$D${13 + i}", name=L.EN, size=9, bg=L.WHITE, fmt=L.INT_RED)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)
        bar = ws.cell(r, 3)
        bar.value = (f'=REPT("█",ROUND(IF(MAX($H$13:$H$23)=0,0,'
                     f'$H{r}/MAX($H$13:$H$23)*20),0))')
        L.st(bar, name=L.EN, size=9, color=L.SAGE, bg=L.WHITE, h="right")
        for c in range(3, 8):
            ws.cell(r, c).fill = fill(L.WHITE)
            ws.cell(r, c).border = box(L.LINE)
    L.note(ws, m0 + 14, last,
           "همه اعداد از دفتر روزنامه (Book) خوانده می‌شوند و با ترازنامه و داشبورد یکی هستند. "
           "برای چاپ گزارش یک گروه، سطر همان گروه را انتخاب کنید.",
           bg=L.MINT, h=30)
    L.page(ws, last, m0 + 14, landscape=False, header_rows="12:12")


# ---------------------------------------------------------------- extras
ENTRY_TIPS = {
    "Scale": {
        "B": "تاریخ شمسی با ارقام لاتین، مثل 1405/07/10. اگر خالی بماند سطر در دفتر نمی‌آید.",
        "C": "ساعت توزین (اختیاری) — برای گزارش شیفت مفید است.",
        "D": "شماره سند باسکول، مثل BS-001. یکتا باشد؛ همین شماره در دفتر روزنامه می‌آید.",
        "E": "نوع عملیات: خرید مرکبات، فروش بار، انعام مصطفی، سایر خروج، هزینه باسکول …",
        "G": "کد حساب طرف مقابل را تایپ/انتخاب کنید: C001 باغدار، B001 بارفروش، BD001 باسکول دیگران.",
        "J": "محصول از فهرست فایل مادر (نارنگی، پرتقال، کیوی …).",
        "K": "نوع سبد؛ روی وزن خالص و مبلغ اثر دارد.",
        "L": "وزن باسکول به کیلوگرم.",
        "M": "تعداد سبد — وزن سبدها از وزن کم می‌شود.",
        "O": "فی خرید (ریال به کیلو).",
        "Q": "مبلغ پز/باسکول برای عملیات پوزی؛ در غیر این صورت خالی.",
        "T": "توضیح آزاد (شماره خودرو، نام راننده، …).",
    },
    "Crate": {
        "B": "تاریخ شمسی 1405/07/10.",
        "D": "نوع عملیات سبد: خرید سبد، فروش سبد، …",
        "E": "کد طرف حساب (S001 فروشنده سبد، C001 باغدار …).",
        "H": "نوع سبد از فهرست (۶ک، ۵ک، …).",
        "I": "تعداد سبد.",
        "J": "قیمت هر سبد (ریال).",
    },
    "Pay": {
        "B": "تاریخ پرداخت/دریافت.",
        "D": "نوع: پرداخت یا دریافت.",
        "E": "روش: نقد، چک، انتقال بانکی، کارت/پوز، پاس چک، برگشت چک، انتقال داخلی.",
        "F": "کد طرف حساب — برای انتقال داخلی، بانک مبدأ را در ستون N بگذارید.",
        "I": "بانک مقصد/عامل: BN01…BN10. برای نقد خالی بماند.",
        "L": "وضعیت چک: در جریان، پاس شده، برگشتی، باطل.",
        "M": "سررسید چک (1405/08/15).",
        "N": "بانک مبدأ برای انتقال داخلی.",
        "P": "توضیح آزاد؛ در دفتر روزنامه می‌آید.",
    },
}


def apply_tips(wb):
    n = 0
    for name, tips in ENTRY_TIPS.items():
        ws = wb[name]
        for col, text in tips.items():
            cell = ws[f"{col}8"]
            cell.comment = Comment(text, "MAZANDASHT", width=300, height=120)
            n += 1
    return n


def extend_help(wb):
    ws = wb["Help"]
    r = ws.max_row + 2
    blocks = [
        ("۱۰", "شیت‌های جدید این نسخه"),
        ("b", "Inv — دفتر فاکتورها: هر ردیف یک قلم فاکتور (خرید، فروش، باسکول، سبد). شماره فاکتور، تاریخ، "
              "نوع عملیات، کد طرف حساب، کالا، تعداد/وزن و فی را وارد کنید؛ مبلغ و مبلغ خالص خودکار است. "
              "ستون «چاپ» را ✔ بزنید تا ردیف در فاکتور چاپی بیاید."),
        ("b", "Print — فاکتور چاپی: فقط ردیف‌های ✔‌خورده را به ترتیب نشان می‌دهد، با جمع کل، مشخصات طرف "
              "حساب و جای مهر/امضا. آماده چاپ A4 عمودی است؛ برای فاکتور بعدی تیک‌های قبلی را بردارید."),
        ("b", "GRep — گزارش گروه‌های حساب: گردش و مانده هر گروه (باغدار، بارفروش، باسکول دیگران، بانک …) "
              "به‌همراه سهم و روند ماهانه؛ آماده چاپ A4."),
        ("b", "ظرفیت: توزین باسکول ۱۰٬۰۰۰ ردیف، سبد ۲٬۵۰۰، پرداخت ۲٬۵۰۰ و دفتر روزنامه ۵۰٬۰۰۰ سطر. "
              "همه فرمول‌ها تا آخرین سطر باز هستند؛ کافی است در سطر خالی بعدی ثبت کنید."),
    ]
    for kind, text in blocks:
        if kind == "۱۰":
            ws.merge_cells(f"A{r}:J{r}")
            c = ws.cell(r, 1, "۱۰   " + text)
            L.st(c, name=L.FA, size=11, bold=True, color=L.WHITE, bg=L.FOREST, border=L.FOREST,
                 h="right")
            ws.row_dimensions[r].height = 26
        else:
            ws.merge_cells(f"A{r}:J{r}")
            c = ws.cell(r, 1, "•  " + text)
            L.st(c, name=L.FA, size=10, color=L.INK, bg=L.MINT, border=L.LINE_S, h="right",
                 wrap=True)
            ws.row_dimensions[r].height = 46
        r += 1
    ws.print_area = f"A1:J{r}"
    return r


def home_links(wb):
    ws = wb["Home"]
    r = ws.max_row + 2
    ws.merge_cells(f"A{r}:J{r}")
    c = ws.cell(r, 1, "   شیت‌های تازه: Inv (دفتر فاکتورها)  ·  Bill (فاکتور خودکار فروش)  ·  "
                      "BillBuy (فاکتور خودکار خرید از باغدار)  ·  Print (فاکتور چاپی A4)  ·  "
                      "GRep (گزارش گروه‌های حساب) — روی نام شیت کلیک کنید.")
    L.st(c, name=L.FA, size=10.5, color=L.WHITE, bg=L.GOLD, border=L.GOLD, h="right")
    ws.row_dimensions[r].height = 24
    links = [("Inv", "دفتر فاکتورها", "B"), ("Bill", "فاکتور خودکار فروش", "C"),
             ("BillBuy", "فاکتور خودکار خرید", "D"),
             ("Print", "فاکتور چاپی", "F"), ("GRep", "گزارش گروه‌ها", "H")]
    for i, (sheet, label, _) in enumerate(links):
        cell = ws.cell(r + 1, 1 + i * 3)
        cell.value = label
        cell.hyperlink = f"#'{sheet}'!A1"
        L.st(cell, name=L.FA, size=10, bold=True, color=L.FOREST, bg=L.MINT, border=L.LINE)
    ws.print_area = f"A1:J{r + 1}"
    return r


def build(path, groups=None):
    wb = openpyxl.load_workbook(path)
    grow.grow_entry(wb)
    n = grow.reseat_ranges(wb, grow.ENTRY_MAP)

    dv_resize(wb["Scale"], {309: grow.SCALE_LAST})
    dv_resize(wb["Crate"], {159: grow.CRATE_LAST})
    dv_resize(wb["Pay"], {159: grow.PAY_LAST})

    import shrk
    shrk.renumber(wb["Scale"], grow.SCALE_LAST)
    shrk.renumber(wb["Pay"], grow.PAY_LAST)
    shrk.renumber(wb["Crate"], grow.CRATE_LAST)
    tips = apply_tips(wb)
    add_inv(wb)
    add_print(wb)
    add_grep(wb, groups, book_last=grow.BOOK_LAST)
    help_rows = extend_help(wb)
    home_links(wb)
    # print areas must follow the new capacities (the v03 files stopped at 309/159)
    wb["Scale"].print_area = f"A1:T{grow.SCALE_LAST}"
    wb["Crate"].print_area = f"A1:M{grow.CRATE_LAST}"
    wb["Pay"].print_area = f"A1:R{grow.PAY_LAST}"
    for ws in wb.worksheets:                 # POS is built by openpyxl: A4 centering
        ws.print_options.horizontalCentered = True

    wb.calculation.fullCalcOnLoad = True
    wb.active = 0
    wb.save(path)
    return {"reseat": n, "tips": tips, "help_rows": help_rows}
