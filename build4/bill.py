# -*- coding: utf-8 -*-
"""Automatic, print-ready invoices (sheets ``Bill`` and ``BillBuy``) for 01_Entry.

Type a customer code in one cell (outside the print area) and the whole
invoice builds itself from the Scale register:

    net kg = FLOOR( (scale kg − basket count × basket tare) × (1 − loss%) , 5 )

* fruit and basket stay dropdowns (data validation) on every line
* ``درصد افت`` and ``ضایعات کل`` live outside the print area, so they can be
  changed without touching the printed page
* a payment block lists the deposits this customer has already made
  (cheque / transfer / cash) straight from the Pay sheet

``build(path)`` makes the **sale** invoice (Bill), ``build(path, "buy")`` makes
the **purchase** invoice (BillBuy — the load received from a grower), with the
same engine: only the filter, the wording and the payment direction change.
"""
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation

import grow
import look as L

SHEET = "Bill"
ROWS = 20            # invoice lines
PAY_ROWS = 10        # payment lines
SCALE_LAST = grow.SCALE_LAST
PAY_LAST = grow.PAY_LAST

C = L          # style shortcuts
WIDTHS = {"A": 6, "B": 11, "C": 22, "D": 15, "E": 8, "F": 12, "G": 11,
          "H": 12, "I": 10, "J": 15, "K": 13}


CFG = {
    "sale": {
        "sheet": "Bill",
        "head": "   فاکتور خودکار فروش — کد مشتری را در کادر زرد (بیرون از چاپ) بزنید؛ "
                "بار، سبد، وزن خالص، افت و واریزی‌ها خودکار می‌آید",
        "party": "کد مشتری", "party_name": "نام مشتری", "party_dv": "ListAllCodes",
        "op": "",                       # no operation filter: every load of the customer
        "pay_type": "دریافت",
        "pay_title": "   واریزی‌های انجام‌شده این مشتری — چک، انتقال بانکی و نقد (از شیت پرداخت)",
        "pay_total": "جمع واریزی این مشتری",
        "balance": "مانده فاکتور",
        "sign_r": "مهر و امضای فروشنده",
        "sign_l": "امضای خریدار",
        "inv_no": "INV-1405-001",
        "loss": 0.01,
        "hdr": "فاکتور فروش",
    },
    "buy": {
        "sheet": "BillBuy",
        "head": "   فاکتور خودکار خرید — کد باغدار را در کادر زرد (بیرون از چاپ) بزنید؛ "
                "بار تحویلی، سبد، وزن خالص و پرداختی‌های ما خودکار می‌آید",
        "party": "کد باغدار", "party_name": "نام باغدار", "party_dv": "ListAllCodes",
        "op": "خرید",
        "pay_type": "پرداخت",
        "pay_title": "   پرداختی‌های انجام‌شده به این باغدار — چک، انتقال بانکی و نقد "
                     "(از شیت پرداخت)",
        "pay_total": "جمع پرداختی به این باغدار",
        "balance": "مانده بدهی به باغدار",
        "sign_r": "مهر و امضای باغدار (فروشنده)",
        "sign_l": "مهر و امضای خریدار — مازندشت",
        "inv_no": "PUR-1405-001",
        "loss": 0.0,
        "hdr": "فاکتور خرید",
    },
}


def _dvs(ws, cfg):
    """Dropdowns for the manual-override columns."""
    dv_prod = DataValidation(type="list", formula1="=ListProducts", allow_blank=True,
                             showDropDown=False, errorStyle="warning",
                             error="کالا را از فهرست فایل مادر انتخاب کنید.",
                             promptTitle="کالا", prompt="نام میوه را از فهرست انتخاب کنید.")
    dv_bask = DataValidation(type="list", formula1="=ListBaskets", allow_blank=True,
                             showDropDown=False, errorStyle="warning",
                             error="سبد را از فهرست فایل مادر انتخاب کنید.",
                             promptTitle="سبد", prompt="نوع سبد را انتخاب کنید.")
    dv_code = DataValidation(type="list", formula1="=" + cfg["party_dv"], allow_blank=True,
                             showDropDown=False, errorStyle="warning",
                             error="کد باید در فایل مادر تعریف شده باشد.",
                             promptTitle=cfg["party"],
                             prompt=cfg["party"] + " را از فهرست انتخاب کنید.")
    for dv in (dv_prod, dv_bask, dv_code):
        ws.add_data_validation(dv)
    dv_code.add(ws["O2"])
    dv_prod.add(ws["O9"])
    dv_bask.add(ws["O10"])


def build(path, mode="sale", index=None):
    global SCALE_LAST, PAY_LAST
    SCALE_LAST, PAY_LAST = grow.SCALE_LAST, grow.PAY_LAST
    cfg = CFG[mode]
    wb = openpyxl.load_workbook(path)
    if cfg["sheet"] in wb.sheetnames:
        del wb[cfg["sheet"]]
    ws = wb.create_sheet(cfg["sheet"], 6 if index is None else index)

    for col, w in WIDTHS.items():
        ws.column_dimensions[col].width = w
    ws.column_dimensions["L"].hidden = True        # scale source index
    ws.column_dimensions["M"].hidden = True        # pay source index
    for col in ("N", "O"):
        ws.column_dimensions[col].width = 20 if col == "N" else 22

    # ------------------------------------------------------------- branding
    ws.merge_cells("A1:K1")
    c = ws["A1"]
    c.value = "   مازندشت  ·  سورتینگ و بسته‌بندی مرکبات"
    C.st(c, name=L.FA_T, size=15, bold=True, color=L.WHITE, bg=L.FOREST, border=L.FOREST,
         h="right")
    ws.row_dimensions[1].height = 30

    ws.merge_cells("A2:K2")
    c = ws["A2"]
    c.value = cfg["head"]
    C.st(c, name=L.FA, size=9.5, color=L.MUTED, bg=L.CREAM, border=L.CREAM, h="right")
    ws.row_dimensions[2].height = 22

    # ---------------------------------------------------------------- info
    ws.merge_cells("A4:K4")
    c = ws["A4"]
    c.value = "   مشخصات فاکتور"
    C.st(c, name=L.FA, size=10.5, bold=True, color=L.WHITE, bg=L.GOLD, border=L.GOLD, h="right")

    labels = [("A5", "شماره فاکتور"), ("C5", "تاریخ فاکتور"), ("E5", cfg["party"]),
              ("G5", cfg["party_name"]), ("I5", "گروه / نوع")]
    for ref, text in labels:
        C.put(ws, ref, text, name=L.FA, size=8.5, bold=True, color=L.MUTED, bg=L.SAND2,
              border=L.LINE, h="center")

    ws.merge_cells("A6:B6")
    ws["A6"] = "=$O$7"
    ws.merge_cells("C6:D6")
    ws["C6"] = "=$O$8"
    ws["E6"] = "=$O$2"
    ws.merge_cells("G6:H6")
    ws["G6"] = '=IF($O$2="","",IFERROR(VLOOKUP($O$2,Sync!$B$10:$G$259,2,0),"کد ناشناخته"))'
    ws.merge_cells("I6:K6")
    ws["I6"] = '=IF($O$2="","",IFERROR(VLOOKUP($O$2,Sync!$B$10:$G$259,3,0),""))'
    for ref in ("A6", "C6", "E6", "G6", "I6"):
        C.st(ws[ref], name=L.FA, size=10.5, bold=True, color=L.FOREST, bg=L.MINT,
             border=L.LINE, h="center")

    # --------------------------------------------------------------- lines
    ws.merge_cells("A8:K8")
    c = ws["A8"]
    c.value = "   اقلام بار  —  وزن خالص = (وزن باسکول − وزن سبدها) × (۱ − درصد افت)  و گرد به پایین مضرب ۵ کیلو"
    C.st(c, name=L.FA, size=10.5, bold=True, color=L.WHITE, bg=L.FOREST, border=L.FOREST,
         h="right")

    heads = [("ردیف", "No."), ("تاریخ", "Date"), ("کالا / میوه", "Product"),
             ("سبد", "Basket"), ("تعداد", "Qty"), ("وزن باسکول", "Gross kg"),
             ("وزن سبد", "Tare kg"), ("وزن خالص", "Net kg"), ("فی", "Price"),
             ("مبلغ", "Amount"), ("سند", "Doc")]
    for i, (fa, en) in enumerate(heads, start=1):
        c = ws.cell(10, i, f"{fa}\n{en}")
        C.st(c, name=L.FA, size=9, bold=True, color=L.WHITE, bg=L.FOREST_2,
             border=L.LINE, h="center", wrap=True)
    ws.row_dimensions[10].height = 32

    first = 11
    last = first + ROWS - 1
    for r in range(first, last + 1):
        n = r - 10
        opf = (f'*ISNUMBER(SEARCH("{cfg["op"]}",Scale!$E$10:$E${SCALE_LAST}))'
               if cfg["op"] else '')
        cond = (f'(((Scale!$G$10:$G${SCALE_LAST}=$O$2))'
                f'*((($O$3="")+(Scale!$B$10:$B${SCALE_LAST}>=$O$3))>0)'
                f'*((($O$4="")+(Scale!$B$10:$B${SCALE_LAST}<=$O$4))>0)'
                f'*((Scale!$E$10:$E${SCALE_LAST}<>""){opf}))')
        ws.cell(r, 12).value = (f'=IFERROR(AGGREGATE(15,6,(ROW(Scale!$B$10:$B${SCALE_LAST}))/'
                                f'{cond},{n}),"")')
        idx = f"$L{r}"
        pos = f"({idx}-9)"
        ws.cell(r, 1).value = f'=IF($B{r}="","",COUNT($B${first}:$B{r}))'
        ws.cell(r, 2).value = f'=IF({idx}="","",INDEX(Scale!$B$10:$B${SCALE_LAST},{pos}))'
        ws.cell(r, 3).value = (f'=IF({idx}="","",IF($O$9<>"",$O$9,'
                                f'INDEX(Scale!$J$10:$J${SCALE_LAST},{pos})))')
        ws.cell(r, 4).value = (f'=IF({idx}="","",IF($O$10<>"",$O$10,'
                                f'INDEX(Scale!$K$10:$K${SCALE_LAST},{pos})))')
        ws.cell(r, 5).value = f'=IF({idx}="","",INDEX(Scale!$M$10:$M${SCALE_LAST},{pos}))'
        ws.cell(r, 6).value = f'=IF({idx}="","",INDEX(Scale!$L$10:$L${SCALE_LAST},{pos}))'
        ws.cell(r, 7).value = (f'=IF(OR($D{r}="",$E{r}=""),"",'
                               f'$E{r}*IFERROR(VLOOKUP($D{r},BasketTbl,4,0),0))')
        ws.cell(r, 8).value = (f'=IF(OR($F{r}="",$G{r}=""),"",'
                               f'MAX(0,FLOOR(($F{r}-$G{r})*(1-$O$5),5)))')
        ws.cell(r, 9).value = f'=IF({idx}="","",INDEX(Scale!$O$10:$O${SCALE_LAST},{pos}))'
        ws.cell(r, 10).value = f'=IF(OR($H{r}="",$I{r}=""),"",$H{r}*$I{r})'
        ws.cell(r, 11).value = f'=IF({idx}="","",INDEX(Scale!$D$10:$D${SCALE_LAST},{pos}))'

        for col in range(1, 12):
            cell = ws.cell(r, col)
            if col in (3, 4):
                C.st(cell, name=L.FA, size=9, color=L.INK, bg=L.INPUT, border=L.LINE_S,
                     h="right")
            elif col in (1, 2, 5, 11):
                C.st(cell, name=L.EN if col != 2 else L.FA, size=9, color=L.INK,
                     bg=L.CREAM, border=L.LINE_S, h="center")
            elif col in (6, 7, 8):
                C.st(cell, name=L.EN, size=9, color=L.INK, bg=L.WHITE, border=L.LINE_S,
                     h="center", fmt=L.INT)
            elif col == 9:
                C.st(cell, name=L.EN, size=9, color=L.INK, bg=L.WHITE, border=L.LINE_S,
                     h="center", fmt=L.INT)
            else:
                C.st(cell, name=L.EN, size=9.5, bold=True, color=L.FOREST, bg=L.MINT,
                     border=L.LINE_S, h="center", fmt=L.INT)
        ws.row_dimensions[r].height = 19

    # -------------------------------------------------------------- totals
    t1, t2, t3 = last + 1, last + 2, last + 3      # 31 / 32 / 33
    ws.merge_cells(f"A{t1}:E{t1}")
    c = ws[f"A{t1}"]
    c.value = "جمع اقلام"
    C.st(c, name=L.FA, size=10, bold=True, color=L.WHITE, bg=L.GOLD, border=L.GOLD, h="center")
    for col, letter in ((6, "F"), (7, "G"), (8, "H"), (10, "J")):
        c = ws.cell(t1, col, f"=SUM({letter}{first}:{letter}{last})")
        C.st(c, name=L.EN, size=10, bold=True, color=L.WHITE, bg=L.GOLD, border=L.GOLD,
             h="center", fmt=L.INT)
    c = ws.cell(t1, 11, "")
    C.st(c, name=L.EN, size=10, color=L.WHITE, bg=L.GOLD, border=L.GOLD, h="center")

    ws.merge_cells(f"A{t2}:G{t2}")
    c = ws[f"A{t2}"]
    c.value = "کسر ضایعات کل (از کادر بیرون چاپ — دستی)"
    C.st(c, name=L.FA, size=9.5, bold=True, color=L.TERRA, bg=L.BLUSH, border=L.LINE,
         h="right")
    c = ws.cell(t2, 8, "=$O$6")
    C.st(c, name=L.EN, size=9.5, bold=True, color=L.TERRA, bg=L.BLUSH, border=L.LINE,
         h="center", fmt=L.INT)
    ws.merge_cells(f"I{t2}:J{t2}")
    c = ws[f"I{t2}"]
    c.value = f'=IFERROR($O$6*($J{t1}/$H{t1}),0)'
    C.st(c, name=L.EN, size=9.5, bold=True, color=L.TERRA, bg=L.BLUSH, border=L.LINE,
         h="center", fmt=L.INT)
    C.st(ws.cell(t2, 11, ""), name=L.FA, size=9, bg=L.BLUSH, border=L.LINE)

    ws.merge_cells(f"A{t3}:G{t3}")
    c = ws[f"A{t3}"]
    c.value = "خالص قابل فروش پس از ضایعات"
    C.st(c, name=L.FA, size=10.5, bold=True, color=L.WHITE, bg=L.FOREST, border=L.FOREST,
         h="right")
    c = ws.cell(t3, 8, f"=$H{t1}-$H{t2}")
    C.st(c, name=L.EN, size=10.5, bold=True, color=L.WHITE, bg=L.FOREST, border=L.FOREST,
         h="center", fmt=L.INT)
    ws.merge_cells(f"I{t3}:J{t3}")
    c = ws[f"I{t3}"]
    c.value = f"=$J{t1}-$I{t2}"
    C.st(c, name=L.EN, size=11, bold=True, color=L.WHITE, bg=L.FOREST, border=L.FOREST,
         h="center", fmt=L.INT)
    C.st(ws.cell(t3, 11, ""), name=L.FA, size=9, bg=L.FOREST, border=L.FOREST)

    # ------------------------------------------------------------ payments
    p0 = t3 + 2                       # 35
    ws.merge_cells(f"A{p0}:K{p0}")
    c = ws[f"A{p0}"]
    c.value = cfg["pay_title"]
    C.st(c, name=L.FA, size=10.5, bold=True, color=L.WHITE, bg=L.SKY_D, border=L.SKY_D,
         h="right")
    ph = ["تاریخ", "روش", "شماره چک", "سررسید", "مبلغ", "وضعیت"]
    for i, h in enumerate(ph, start=1):
        c = ws.cell(p0 + 1, i, h)
        C.st(c, name=L.FA, size=9, bold=True, color=L.WHITE, bg=L.SKY_D, border=L.LINE,
             h="center")
    ws.cell(p0 + 1, 7).value = "بانک / توضیح"
    ws.merge_cells(start_row=p0 + 1, start_column=7, end_row=p0 + 1, end_column=11)
    C.st(ws.cell(p0 + 1, 7), name=L.FA, size=9, bold=True, color=L.WHITE, bg=L.SKY_D,
         border=L.LINE, h="center")

    pf = p0 + 2                              # first payment row
    for i in range(PAY_ROWS):
        r = pf + i
        n = i + 1
        ws.cell(r, 13).value = (
            f'=IFERROR(AGGREGATE(15,6,(ROW(Pay!$B$10:$B${PAY_LAST}))/'
            f'((Pay!$F$10:$F${PAY_LAST}=$O$2)'
            f'*(Pay!$D$10:$D${PAY_LAST}="{cfg["pay_type"]}")),{n}),"")')
        idx = f"$M{r}"
        pos = f"({idx}-9)"
        ws.cell(r, 1).value = f'=IF({idx}="","",INDEX(Pay!$B$10:$B${PAY_LAST},{pos}))'
        ws.cell(r, 2).value = f'=IF({idx}="","",INDEX(Pay!$E$10:$E${PAY_LAST},{pos}))'
        ws.cell(r, 3).value = f'=IF({idx}="","",INDEX(Pay!$J$10:$J${PAY_LAST},{pos}))'
        ws.cell(r, 4).value = f'=IF({idx}="","",INDEX(Pay!$M$10:$M${PAY_LAST},{pos}))'
        ws.cell(r, 5).value = f'=IF({idx}="","",INDEX(Pay!$K$10:$K${PAY_LAST},{pos}))'
        ws.cell(r, 6).value = f'=IF({idx}="","",INDEX(Pay!$L$10:$L${PAY_LAST},{pos}))'
        ws.cell(r, 7).value = f'=IF({idx}="","",INDEX(Pay!$I$10:$I${PAY_LAST},{pos}))'
        ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=11)
        for col in range(1, 12):
            cell = ws.cell(r, col)
            if col == 5:
                C.st(cell, name=L.EN, size=9.5, bold=True, color=L.SKY_D, bg=L.SKY,
                     border=L.LINE_S, h="center", fmt=L.INT)
            elif col in (2, 6):
                C.st(cell, name=L.FA, size=9, color=L.INK, bg=L.WHITE, border=L.LINE_S,
                     h="center")
            else:
                C.st(cell, name=L.EN if col in (1, 3, 4) else L.FA, size=9, color=L.INK,
                     bg=L.WHITE, border=L.LINE_S, h="center")
        ws.row_dimensions[r].height = 19

    pt1 = pf + PAY_ROWS                          # totals row
    ws.merge_cells(f"A{pt1}:C{pt1}")
    c = ws[f"A{pt1}"]
    c.value = cfg["pay_total"]
    C.st(c, name=L.FA, size=10, bold=True, color=L.WHITE, bg=L.SKY_D, border=L.SKY_D,
         h="right")
    c = ws.cell(pt1, 4, f'=SUMIFS(Pay!$K$10:$K${PAY_LAST},Pay!$F$10:$F${PAY_LAST},$O$2,'
                        f'Pay!$D$10:$D${PAY_LAST},"{cfg["pay_type"]}")')
    C.st(c, name=L.EN, size=10.5, bold=True, color=L.WHITE, bg=L.SKY_D, border=L.SKY_D,
         h="center", fmt=L.INT)
    ws.cell(pt1, 5).value = cfg["balance"]
    C.st(ws.cell(pt1, 5), name=L.FA, size=9, bold=True, color=L.WHITE, bg=L.SKY_D,
         border=L.SKY_D, h="center")
    ws.merge_cells(start_row=pt1, start_column=6, end_row=pt1, end_column=11)
    c = ws.cell(pt1, 6, f"=$I{t3}-$D{pt1}")
    C.st(c, name=L.EN, size=11, bold=True, color=L.WHITE, bg=L.SKY_D, border=L.SKY_D,
         h="center", fmt=L.INT)

    # ----------------------------------------------------------- signatures
    sr = pt1 + 3
    ws.merge_cells(start_row=sr, start_column=1, end_row=sr, end_column=5)
    ws.cell(sr, 1).value = cfg["sign_r"]
    ws.merge_cells(start_row=sr, start_column=7, end_row=sr, end_column=11)
    ws.cell(sr, 7).value = cfg["sign_l"]
    for ref, txt in ((ws.cell(sr, 1), cfg["sign_r"]), (ws.cell(sr, 7), cfg["sign_l"])):
        C.st(ref, name=L.FA, size=9.5, bold=True, color=L.MUTED, bg=L.CREAM,
             border=L.LINE, h="center")
    ws.row_dimensions[sr].height = 44

    fr = sr + 1
    ws.merge_cells(start_row=fr, start_column=1, end_row=fr, end_column=11)
    c = ws.cell(fr, 1)
    c.value = ("برای چاپ: درصد افت و ضایعات کل را در کادرهای بیرون چاپ (N/O) تنظیم کنید؛ "
               "همین صفحه با Ctrl+P آماده چاپ A4 است.")
    C.st(c, name=L.FA, size=8.5, color=L.MUTED, bg=L.SAND, border=L.SAND, h="center")

    # ------------------------------------------------------------ controls
    ws.merge_cells("N1:O1")
    c = ws["N1"]
    c.value = "کادرهای کنترل — بیرون از محدوده چاپ"
    C.st(c, name=L.FA, size=10, bold=True, color=L.WHITE, bg=L.TERRA, border=L.TERRA,
         h="center")

    ctrl = [("N2", cfg["party"] + "  *", "O2", None, L.INPUT),
            ("N3", "از تاریخ (اختیاری)", "O3", None, L.INPUT),
            ("N4", "تا تاریخ (اختیاری)", "O4", None, L.INPUT),
            ("N5", "درصد افت (مثلاً 1%)", "O5", cfg["loss"], L.INPUT),
            ("N6", "ضایعات کل (kg)", "O6", 0, L.INPUT),
            ("N7", "شماره فاکتور", "O7", cfg["inv_no"], L.INPUT),
            ("N8", "تاریخ فاکتور", "O8", "1405/07/10", L.INPUT)]
    for lab_ref, lab, val_ref, default, bg in ctrl:
        ws[lab_ref] = lab
        C.st(ws[lab_ref], name=L.FA, size=9, bold=True, color=L.INK, bg=L.SAND2,
             border=L.LINE, h="right")
        cell = ws[val_ref]
        if default is not None:
            cell.value = default
        C.st(cell, name=L.EN if val_ref in ("O2", "O3", "O4", "O5", "O6") else L.FA,
             size=10.5, bold=True, color=L.FOREST, bg=bg, border=L.GOLD, h="center")
        if val_ref == "O5":
            cell.number_format = "0.0%"

    ctrl2 = [("N9", "میوهٔ پیش‌فرض (اختیاری)", "O9", L.INPUT),
             ("N10", "سبد پیش‌فرض (اختیاری)", "O10", L.INPUT)]
    for lab_ref, lab, val_ref, bg in ctrl2:
        ws[lab_ref] = lab
        C.st(ws[lab_ref], name=L.FA, size=9, bold=True, color=L.INK, bg=L.SAND2,
             border=L.LINE, h="right")
        C.st(ws[val_ref], name=L.FA, size=9.5, bold=True, color=L.FOREST, bg=bg,
             border=L.GOLD, h="center")

    ws["N12"] = "راهنما"
    C.st(ws["N12"], name=L.FA, size=9, bold=True, color=L.WHITE, bg=L.FOREST,
         border=L.FOREST, h="center")
    ws.merge_cells("N13:O22")
    c = ws["N13"]
    c.value = ("۱) کد مشتری را در O2 انتخاب کنید؛ همه ردیف‌ها خودکار می‌آید.\n"
               "۲) اگر بازه تاریخ بدهید، فقط بار همان بازه می‌آید.\n"
               "۳) درصد افت روی همه ردیف‌ها یک‌جا اثر می‌گذارد.\n"
               "۴) ضایعات کل (کیلو) از جمع کم می‌شود و مبلغش هم محاسبه می‌شود.\n"
               "۵) نام میوه و سبد هر ردیف خودکار از باسکول می‌آید؛ برای یکدست کردن "
               "کل فاکتور، O9 و O10 را از فهرست کشویی انتخاب کنید.\n"
               "۶) هیچ‌کدام از این کادرها چاپ نمی‌شوند.")
    C.st(c, name=L.FA, size=8.5, color=L.INK, bg=L.CREAM, border=L.LINE, h="right", wrap=True)

    _dvs(ws, cfg)

    ws.print_area = f"A1:K{fr}"
    ws.page_setup.paperSize = 9
    ws.page_setup.orientation = "portrait"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_options.horizontalCentered = True
    ws.oddHeader.left.text = "مازندشت — " + cfg["hdr"]
    ws.oddHeader.center.text = ("MAZANDASHT · " +
                                ("PURCHASE INVOICE" if mode == "buy" else "SALES INVOICE"))
    ws.oddHeader.right.text = "&A"
    ws.oddFooter.left.text = "&F"
    ws.oddFooter.center.text = "صفحه &P از &N"
    ws.oddFooter.right.text = "نسخه ۰۴ · ۱۴۰۵"
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False

    wb.save(path)
    return ws.max_row
