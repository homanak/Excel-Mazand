# -*- coding: utf-8 -*-
"""v04.6 — تقسیم مبلغ، کارمزد بانکی و صورت‌حساب.

Three things the client asked for that all live in the same block:

1. one payment/receipt can be **split** between up to three other accounts
   (the block sits outside the print area, exactly as asked);
2. the **bank fee** is written in its own column, so the counterpart account
   keeps the clean amount while the bank leg carries amount ± fee;
3. every journal line keeps the document number, so a **statement** per account
   (in / out, stacked, printable) can be built from `Book`.

The block writes four extra legs per payment row, so `grow` now rebuilds the
journal with 6 lines per payment (see `grow.pay_legs`).
"""
import re

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

import grow
import look as L

FA, EN = L.FA, L.EN

FEE_CODE = "X022"
FEE_NAME = "کارمزد و هزینه‌های بانکی"
ACC_FEE_ROW = 281          # 00_Master ← Acc
SYNC_FEE_ROW = 280         # 01_Entry ← Sync (accounts block 10…280)
SYNC_LAST = 280
ACC_NAMES = ("AcctTbl", "AcctCodes", "ListAllCodes", "ListAcctNames")

# (column, fa header, en header, kind, tooltip)
BLOCK = [
    ("AH", "کارمزد", "Fee", "money",
     "کارمزد بانکی همین تسویه. اگر صورتحساب بانک ۴۵٬۰۰۲٬۴۰۰ است، مبلغ را ۴۵٬۰۰۰٬۰۰۰ و "
     "کارمزد را ۲٬۴۰۰ بنویسید؛ حساب طرف حساب پاک می‌ماند و کارمزد به حساب کارمزد می‌رود."),
    ("AI", "حساب کارمزد", "Fee acct", "acct",
     "حسابی که کارمزد به آن می‌خورد (پیش‌فرض X022 — کارمزد و هزینه‌های بانکی)."),
    ("AJ", "حساب تقسیم ۱", "Split 1 acct", "acct",
     "اگر همین مبلغ بین چند حساب (خرید، کرایه، هزینه …) تقسیم می‌شود، اینجا حساب اول را بزنید."),
    ("AK", "مبلغ تقسیم ۱", "Split 1", "money",
     "سهم حساب اول. این مبلغ از حساب اصلی (کد حساب ستون F) کم می‌شود و خودش یک سطر دفتر می‌سازد."),
    ("AL", "حساب تقسیم ۲", "Split 2 acct", "acct",
     "حساب دوم برای تقسیم همین تسویه (اختیاری)."),
    ("AM", "مبلغ تقسیم ۲", "Split 2", "money", "سهم حساب دوم."),
    ("AN", "حساب تقسیم ۳", "Split 3 acct", "acct",
     "حساب سوم برای تقسیم همین تسویه (اختیاری)."),
    ("AO", "مبلغ تقسیم ۳", "Split 3", "money", "سهم حساب سوم."),
    ("AP", "کنترل تقسیم", "Check", "calc",
     "✓ یعنی تقسیم و کارمزد درست است · ⚠ یعنی جمع تقسیم‌ها از مبلغ بیشتر است یا در این روش "
     "پرداخت مجاز نیست (انتقال داخلی، پاس چک، برگشت چک)."),
]


def _box(color="FFDDD5C4"):
    from openpyxl.styles import Border, Side
    s = Side(style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)


def _def(wb, name, ref):
    try:
        del wb.defined_names[name]
    except KeyError:
        pass
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def _dvs(ws, ref, spec):
    dv = DataValidation(type=spec["validate"],
                        formula1=("=" + spec["source"]) if spec.get("source") else None)
    dv.allow_blank = True
    if spec.get("prompt"):
        dv.promptTitle = spec.get("promptTitle", "")
        dv.prompt = spec["prompt"]
        dv.showInputMessage = True
    if spec.get("error"):
        dv.errorTitle = spec.get("errorTitle", "")
        dv.error = spec["error"]
        dv.showErrorMessage = True
    ws.add_data_validation(dv)
    dv.add(ref)
    return dv


# ------------------------------------------------------------------ 00_Master
def master(path):
    """Add the bank-fee account and widen the account tables."""
    wb = openpyxl.load_workbook(path)
    ws = wb["Acc"]
    used = [r for r in range(10, 340) if ws.cell(r, 1).value]
    codes = {str(ws.cell(r, 1).value) for r in used}
    made = 0
    if FEE_CODE not in codes:
        r = used[-1] + 1
        vals = [FEE_CODE, "سایر حساب", "هزینه", FEE_NAME, "—", "—", "—", 0,
                "حساب هزینه", "بله", f'=IF($A{r}="","",IF(LEFT($A{r},2)="X0","سایر حساب",""))',
                "همه کارمزدها و هزینه‌های بانکی اینجا جمع می‌شود."]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(r, c, v)
            cell._style = ws.cell(used[-1], c)._style
        ws.print_area = f"A1:M{r + 1}"
        made = 1
    n = _extend(wb, used[-1] + made)
    wb.save(path)
    wb.close()
    return {"acc": made, "names": n}


def _extend(wb, last):
    """Account tables 10…259 / 279 / 280 → 10…last."""
    n = 0
    for name, dn in list(wb.defined_names.items()):
        if name not in ACC_NAMES:
            continue
        txt = dn.attr_text
        new = re.sub(r"\$([A-Z]{1,3})\$10:\$([A-Z]{1,3})\$\d+",
                     lambda m: f"${m.group(1)}$10:${m.group(2)}${last}", txt)
        new = re.sub(r"\$C\$\d+", f"$C${last}", new)
        if new != txt:
            dn.attr_text = new
            n += 1
    return n


# ------------------------------------------------------------------- 01_Entry
def sync(wb):
    """One more row in the Sync account block: the bank-fee account."""
    ws = wb["Sync"]
    vals = [FEE_CODE, FEE_NAME, "سایر حساب", "هزینه", "حساب هزینه", 0]
    for c, v in enumerate(vals, start=2):
        cell = ws.cell(SYNC_FEE_ROW, c, v)
        try:
            cell._style = ws.cell(SYNC_FEE_ROW - 1, c)._style
        except Exception:
            pass
    for name, dn in list(wb.defined_names.items()):
        if name not in ACC_NAMES:
            continue
        txt = dn.attr_text
        new = re.sub(r"\$([A-Z]{1,3})\$10:\$([A-Z]{1,3})\$\d+",
                     lambda m: f"${m.group(1)}$10:${m.group(2)}${SYNC_LAST}", txt)
        new = re.sub(r"\$C\$\d+", f"$C${SYNC_LAST}", new)
        if new != txt:
            dn.attr_text = new
    return 1


def pay(wb):
    """The out-of-print split / fee block on the payment sheet."""
    ws = wb["Pay"]
    last = grow.PAY_LAST
    hdr7 = ws["R7"]._style if ws["R7"].value is None else ws["R7"]._style
    fa8, en9 = ws["R8"]._style, ws["R9"]._style
    head7 = ws["Q5"]._style

    # band
    ws.merge_cells("AH7:AI7")
    ws.merge_cells("AJ7:AO7")
    for ref, text in (("AH7", "کارمزد بانکی  ·  BANK FEE"),
                      ("AJ7", "تقسیم مبلغ بین حساب‌های دیگر  ·  SPLIT"),
                      ("AP7", "کنترل")):
        c = ws[ref]
        c.value = text
        c._style = hdr7
        c.fill = PatternFill("solid", fgColor="FFF3EADB")
        c.font = Font(name=FA, size=9, bold=True, color="FF14453D")
        c.alignment = Alignment(horizontal="center", vertical="center", readingOrder=2)

    for col, fa, en, kind, tip in BLOCK:
        for row, text, style in ((8, fa, fa8), (9, en, en9)):
            c = ws[f"{col}{row}"]
            c.value = text
            c._style = style
        c = ws[f"{col}8"]
        try:
            from openpyxl.comments import Comment
            c.comment = Comment(tip, "MAZANDASHT", width=320, height=140)
        except Exception:
            pass
        ws.column_dimensions[col].width = 13 if kind != "acct" else 12
        ws[f"{col}7"].fill = PatternFill("solid", fgColor="FFF3EADB")

    # input / calc cells
    f = {
        "AP": ('=IF($D{r}="","",IF(AND($AQ{r}=0,N($AH{r})=0),"",'
               'IF(AND($AQ{r}>0,$AS{r}=0),"⚠ این روش تقسیم ندارد",'
               'IF(AND($AQ{r}>0,$AK{r}+$AM{r}+$AO{r}=0),"⚠ مبلغ تقسیم خالی",'
               'IF($AQ{r}+N($AH{r})>$K{r},"⚠ بیشتر از مبلغ",'
               'IF(OR(AND($AK{r}>0,$AJ{r}=""),AND($AM{r}>0,$AL{r}=""),'
               'AND($AO{r}>0,$AN{r}="")),"⚠ حساب تقسیم خالی","✓"))))))'),
        "AQ": '=SUM($AK{r},$AM{r},$AO{r})',
        "AR": ('=IF($K{r}="","",$K{r}+IF($D{r}="پرداخت",N($AH{r}),'
               'IF($D{r}="دریافت",-N($AH{r}),0)))'),
        "AS": ('=IF(OR($E{r}="پاس چک",$E{r}="برگشت چک",$E{r}="انتقال داخلی"),0,$AQ{r})'),
        "AT": ('=IF($D{r}="","",IF($K{r}="",0,MAX(0,$K{r}-IF($D{r}="پرداخت",$AS{r},'
               'IF($D{r}="دریافت",N($AH{r}),0)))))'),
        "AU": ('=IF($D{r}="","",IF($K{r}="",0,$K{r}+IF($D{r}="پرداخت",N($AH{r}),0)'
               '-IF($D{r}="دریافت",$AS{r},0)))'),
    }
    for r in range(10, last + 1):
        for col, tpl in f.items():
            ws[f"{col}{r}"].value = tpl.format(r=r)
        for col, _, _, kind, _ in BLOCK:
            if kind == "acct":
                ws[f"{col}{r}"].number_format = "General"
        ws[f"AI{r}"].value = FEE_CODE

    # styles: inputs yellow, calcs mint
    for col, _, _, kind, _ in BLOCK:
        for r in range(10, last + 1):
            c = ws[f"{col}{r}"]
            c.border = _box()
            c.alignment = Alignment(horizontal="center", vertical="center",
                                    readingOrder=2)
            c.font = Font(name=EN if kind != "acct" else FA, size=9, color="FF2B2A26")
            if kind == "money":
                c.number_format = "#,##0"
                c.fill = PatternFill("solid", fgColor="FFFFF4D6")
            elif kind == "acct":
                c.fill = PatternFill("solid", fgColor="FFFFF4D6")
            else:
                c.fill = PatternFill("solid", fgColor="FFE4F0E8")
    for col in ("AQ", "AR", "AS", "AT", "AU"):
        for r in range(10, last + 1):
            c = ws[f"{col}{r}"]
            c.fill = PatternFill("solid", fgColor="FFE4F0E8")
            c.font = Font(name=EN, size=8.5, color="FF5E5A52")
            if col in ("AR", "AT", "AU"):
                c.number_format = "#,##0"
    for col in ("AQ", "AR", "AS", "AT", "AU"):
        ws.column_dimensions[col].hidden = True

    _dvs(ws, f"AI10:AI{last}", {"validate": "list", "source": "ListAllCodes",
                                "promptTitle": "حساب کارمزد",
                                "prompt": "حسابی که کارمزد بانکی به آن می‌خورد."})
    for col in ("AJ", "AL", "AN"):
        _dvs(ws, f"{col}10:{col}{last}", {"validate": "list",
                                          "source": "ListAllCodes",
                                          "promptTitle": "حساب تقسیم",
                                          "prompt": "سهم این مبلغ به این حساب می‌رود."})
    ws.merge_cells("AH4:AI4")
    c = ws["AH4"]
    c.value = "خارج از چاپ: کارمزد و تقسیم"
    c.font = Font(name=FA, size=9, bold=True, color="FFC4623F")
    c.alignment = Alignment(horizontal="center", vertical="center", readingOrder=2)
    return len(BLOCK)


def entry(path, master_path):
    """Everything the split/fee block needs inside 01_Entry."""
    wb = openpyxl.load_workbook(path)
    sync(wb)
    pay(wb)
    codes(wb)
    stmt(wb)
    ledger(wb)
    counters(wb)
    book_counter(wb)
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)
    wb.close()
    return {"sync": 1, "sheets": ["Stmt", "Ledger"]}


# ------------------------------------------------------------ codes + counters
def codes(wb):
    """«کد تراکنش» PY-0001… in Pay (out of print, next to the split block)."""
    ws = wb["Pay"]
    for r in range(10, grow.PAY_LAST + 1):
        c = ws[f"AV{r}"]
        c.value = f'=IF($B{r}="","","PY-"&TEXT(N($A{r}),"0000"))'
        c.font = Font(name=EN, size=9, color="FF2E5E4E")
        c.alignment = Alignment(horizontal="center", vertical="center")
    head = ws["AV8"]
    head.value = "کد تراکنش"
    head._style = ws["R8"]._style
    ws["AV9"].value = "Code"
    ws["AV9"]._style = ws["R9"]._style
    ws.column_dimensions["AV"].width = 12
    return 1


def counters(wb):
    """Running counters that turn the two selectors into row lookups."""
    sc = wb["Scale"]
    for r in range(10, grow.SCALE_LAST + 1):
        sc[f"BA{r}"].value = f'=N(BA{r - 1})+IF($U{r}=Stmt!$C$9,1,0)'
    pay = wb["Pay"]
    for r in range(10, grow.PAY_LAST + 1):
        pay[f"AW{r}"].value = f'=N(AW{r - 1})+IF($F{r}=Stmt!$C$9,1,0)'
    for ws, col in ((sc, "BA"), (pay, "AW")):
        ws.column_dimensions[col].hidden = True
    return 1


def book_counter(wb):
    """Book!S — a date key for the account typed in Ledger!C9 (sorted ledger)."""
    bk = wb["Book"]
    last = grow.BOOK_LAST
    for r in range(10, last + 1):
        bk[f"S{r}"].value = (
            f'=IF($D{r}="","",IF($D{r}=Ledger!$C$9,'
            f'IFERROR(VALUE(SUBSTITUTE(LEFT($A{r},10),"/",""))+ROW()/100000,'
            f'90000000+ROW()),""))')
        bk[f"S{r}"].font = Font(name=EN, size=8, color="FF8C8677")
    bk.column_dimensions["S"].hidden = True
    return last - 9


# ------------------------------------------------------------------- Stmt
STMT_COLS = [
    ("تاریخ", "Date", 12), ("سند", "Doc", 12), ("عملیات", "Operation", 20),
    ("محصول", "Product", 14), ("وزن باسکول", "Scale kg", 12), ("تار سبد", "Tare", 11),
    ("وزن خالص", "Net kg", 12), ("فی", "Price", 11), ("مبلغ", "Amount", 16),
    ("جهت", "Dir", 9), ("توضیح", "Note", 24),
]


def _sheet(wb, name, after):
    if name in wb.sheetnames:
        del wb[name]
    idx = wb.sheetnames.index(after) + 1 if after in wb.sheetnames else len(wb.sheetnames)
    return wb.create_sheet(name, idx)


def stmt(wb):
    """Sheet ``Stmt`` — 30 weighbridge receipts + 8 payments + the balance."""
    S, P, B = grow.SCALE_LAST, grow.PAY_LAST, grow.BOOK_LAST
    ws = _sheet(wb, "Stmt", "Pay")
    last = 11
    L.brand(ws, last, "صورتحساب طرف حساب — رسید بارها و دریافتی‌ها",
            "کد حساب را در کادر زرد بزنید؛ بارها از باسکول و پرداخت‌ها از شیت Pay خودکار می‌آید. "
            "آماده چاپ A4.", icon="◈")
    L.kpis(ws, 6, last, [
        ("مانده حساب", f"=$C$8", "#,##0;[Red]-#,##0", L.FOREST),
        ("جمع بدهکار", f"=SUMIFS(Book!$L$10:$L${B},Book!$D$10:$D${B},$C$9)", "#,##0;[Red]-#,##0", L.TERRA),
        ("جمع بستانکار", f"=SUMIFS(Book!$M$10:$M${B},Book!$D$10:$D${B},$C$9)", "#,##0;[Red]-#,##0", L.GOLD),
        ("وزن بار (kg)", f"=SUMIFS(Scale!$N$10:$N${S},Scale!$U$10:$U${S},$C$9)", "#,##0.00", L.SAGE),
    ])
    # the balance itself lives in B8 (below the KPI strip, out of the print flow)
    ws.merge_cells("A8:B8")
    c = ws["A8"]
    c.value = "مانده"
    L.st(c, name=FA, size=9, bold=True, color=L.MUTED, bg=L.SAND2, border=L.LINE)
    ws.merge_cells("C8:E8")
    c = ws["C8"]
    c.value = (f'=SUMIFS(Book!$L$10:$L${B},Book!$D$10:$D${B},$C$9)'
               f'-SUMIFS(Book!$M$10:$M${B},Book!$D$10:$D${B},$C$9)')
    L.st(c, name=EN, size=11, bold=True, color=L.FOREST, bg=L.MINT, border=L.LINE,
         fmt="#,##0;[Red]-#,##0")
    ws.merge_cells("F8:K8")
    c = ws["F8"]
    c.value = (f'=IF($C$9="","",IF(ROUND($C$8,0)>0,"بدهی ایشان به شرکت: "&TEXT($B$8,"#,##0")&" ریال",'
               f'IF(ROUND($C$8,0)<0,"بستانکاری (بدهی شرکت): "&TEXT(-$B$8,"#,##0")&" ریال","تسویه — مانده صفر")))')
    L.st(c, name=FA, size=9.5, bold=True, color=L.FOREST, bg=L.CREAM, border=L.LINE, h="right")

    ws.row_dimensions[9].height = 26
    ws.merge_cells("A9:B9")
    c = ws["A9"]
    c.value = "کد حساب"
    L.st(c, name=FA, size=10, bold=True, color=L.WHITE, bg=L.FOREST, border=L.FOREST)
    c = ws["C9"]
    c.value = "C001"
    L.st(c, name=EN, size=10, bold=True, color=L.INK, bg=L.INPUT, border=L.TERRA)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells("D9:K9")
    c = ws["D9"]
    c.value = ('=IF($C$9="","کد حساب را انتخاب کنید (مثلاً C001 یا SL001 یا SH01) …",'
               'IFERROR(VLOOKUP($C$9,AcctTbl,2,0)&"   ·   "&VLOOKUP($C$9,AcctTbl,3,0)'
               '&"   ·   "&VLOOKUP($C$9,AcctTbl,4,0),"⚠ این کد در جدول حساب‌ها نیست"))')
    L.st(c, name=FA, size=9.5, bold=True, color=L.FOREST, bg=L.SAND, border=L.LINE, h="right")
    _dvs(ws, "C9", {"validate": "list", "source": "ListAllCodes",
                    "promptTitle": "کد حساب", "prompt": "کد باغدار، بارفروش، شریک یا هر حساب دیگر."})

    ws["M2"].value = f'=COUNTIF(Scale!$U$10:$U${S},$C$9)'
    ws["M3"].value = f'=COUNTIF(Pay!$F$10:$F${P},$C$9)'
    for ref in ("M2", "M3"):
        ws[ref].font = Font(name=EN, size=8, color="FF8C8677")

    L.section(ws, 11, last, "۰۱", "رسید بارها — ۳۰ ردیف آخر از باسکول", tint=L.FOREST)
    ws.row_dimensions[11].height = 24
    L.thead(ws, 12, [(c[0], c[1], c[2]) for c in STMT_COLS], height=30)
    for k in range(1, 31):
        r = 12 + k
        ws[f"M{r}"].value = (f"=IF($M$2<{k},0,IF($M$2-30+{k}<1,{k},$M$2-30+{k}))")
        vals = {
            "A": "B", "B": "D", "C": "E", "D": "J", "E": "L", "F": "M",
            "G": "N", "H": "O", "I": "AJ", "J": "F", "K": "T",
        }
        for c, src in vals.items():
            cell = ws[f"{c}{r}"]
            cell.value = (f'=IF($M{r}<1,"",IFERROR(INDEX(Scale!${src}$10:${src}${S},'
                          f'MATCH($M{r},Scale!$BA$10:$BA${S},0)),""))')
            L.st(cell, name=EN if c in "ABEFGHIJ" else FA, size=9, color=L.INK,
                 bg=L.CREAM if k % 2 else L.WHITE, h="center" if c != "K" else "right",
                 fmt=("#,##0.00" if c in "EFG" else ("#,##0" if c in "HI" else None)))
        ws.row_dimensions[r].height = 17

    L.section(ws, 44, last, "۰۲", "دریافتی‌ها و پرداختی‌ها — ۸ ردیف آخر", tint=L.TERRA)
    L.thead(ws, 45, [("تاریخ", "Date", 12), ("سند", "Doc", 12), ("نوع", "Type", 20),
                     ("روش", "Method", 14), ("شماره چک", "Cheque", 12),
                     ("وضعیت چک", "Cheque st.", 11), ("مبلغ", "Amount", 16),
                     ("بدهکار", "Debit", 14), ("بستانکار", "Credit", 14),
                     ("مانده", "Balance", 16), ("توضیح", "Note", 24)], tint=L.TERRA,
                    height=30)
    for k in range(1, 9):
        r = 45 + k
        ws[f"M{r}"].value = (f"=IF($M$3<{k},0,IF($M$3-8+{k}<1,{k},$M$3-8+{k}))")
        vals = {"A": "B", "B": "C", "C": "D", "D": "E", "E": "J", "F": "L",
                "G": "K", "K": "P"}
        for c, src in vals.items():
            cell = ws[f"{c}{r}"]
            cell.value = (f'=IF($M{r}<1,"",IFERROR(INDEX(Pay!${src}$10:${src}${P},'
                          f'MATCH($M{r},Pay!$AW$10:$AW${P},0)),""))')
            L.st(cell, name=EN if c in "ABEFG" else FA, size=9, color=L.INK,
                 bg=L.CREAM if k % 2 else L.WHITE, h="center" if c != "K" else "right")
        cell = ws[f"H{r}"]
        cell.value = f'=IF($C{r}="پرداخت",N($G{r}),"")'
        L.st(cell, name=EN, size=9, color=L.TERRA, bg=L.CREAM if k % 2 else L.WHITE,
             fmt="#,##0;[Red]-#,##0")
        cell = ws[f"I{r}"]
        cell.value = f'=IF($C{r}="دریافت",N($G{r}),"")'
        L.st(cell, name=EN, size=9, color=L.SAGE, bg=L.CREAM if k % 2 else L.WHITE,
             fmt="#,##0;[Red]-#,##0")
        ws.row_dimensions[r].height = 17

    for col, w in (("A", 11), ("B", 11), ("C", 19), ("D", 15), ("E", 12), ("F", 10),
                   ("G", 11), ("H", 11), ("I", 14), ("J", 9), ("K", 22)):
        ws.column_dimensions[col].width = w

    for row, label, formula, fmt, bg in (
            (54, "جمع بار", f'=SUM(I13:I42)', "#,##0", L.MINT),
            (55, "جمع دریافتی / پرداختی", f'=SUM(H46:I53)', "#,##0", L.MINT),
            (56, "مانده حساب", "=$C$8", "#,##0;[Red]-#,##0", L.SAND)):
        ws.merge_cells(f"A{row}:D{row}")
        c = ws[f"A{row}"]
        c.value = label
        L.st(c, name=FA, size=10, bold=True, color=L.FOREST, bg=bg, border=L.LINE, h="right")
        for col in "BCDE":
            ws[f"{col}{row}"].fill = PatternFill("solid", fgColor=bg)
            ws[f"{col}{row}"].border = _box()
        ws.merge_cells(f"E{row}:H{row}")
        c = ws[f"E{row}"]
        c.value = formula
        L.st(c, name=EN, size=11, bold=True, color=L.FOREST, bg=L.CREAM, border=L.LINE,
             fmt=fmt)
        for col in "FGH":
            ws[f"{col}{row}"].fill = PatternFill("solid", fgColor=L.CREAM)
            ws[f"{col}{row}"].border = _box()
        ws.row_dimensions[row].height = 22
    ws.merge_cells("I54:K54")
    c = ws["I54"]
    c.value = "جمع وزن و مبلغ بارهای همین صفحه"
    L.st(c, name=FA, size=8.5, color=L.MUTED, bg=L.CREAM, border=L.LINE, h="right")
    ws.merge_cells("I55:K55")
    c = ws["I55"]
    c.value = "جمع ردیف‌های دریافتی/پرداختی همین صفحه"
    L.st(c, name=FA, size=8.5, color=L.MUTED, bg=L.CREAM, border=L.LINE, h="right")
    ws.merge_cells("I56:K56")
    c = ws["I56"]
    c.value = "مانده کل حساب از دفتر روزنامه (همه ردیف‌ها، نه فقط این صفحه)"
    L.st(c, name=FA, size=8.5, color=L.MUTED, bg=L.SAND, border=L.LINE, h="right")

    L.note(ws, 58, last,
           "این برگ برای دادن به باغدار یا بارفروش است: ۳۰ رسید بار آخر بالا و ۸ دریافتی آخر پایین. "
           "برای تاریخچه کامل همان حساب، شیت Ledger را باز کنید. اگر بارها بدون فی ثبت شده‌اند "
           "(بار ورودی به سورتینگ)، ستون مبلغ صفر می‌ماند و مانده از دفتر روزنامه خوانده می‌شود.",
           bg=L.MINT, h=40)
    L.page(ws, last, 58, landscape=False)
    ws.page_setup.fitToHeight = 1
    ws.print_area = "A1:K58"
    return "Stmt"


def ledger(wb):
    """Sheet ``Ledger`` — every journal line of one account, sorted by date."""
    B = grow.BOOK_LAST
    N = 150
    ws = _sheet(wb, "Ledger", "Stmt")
    L.brand(ws, 8, "دفتر حساب — همه تراکنش‌های یک حساب",
            "کد حساب را بزنید؛ همه ردیف‌های بدهکار و بستانکار آن حساب، به ترتیب تاریخ، "
            "زیر هم و با ماندهِ روان.", icon="◈", tint=L.INK)
    L.kpis(ws, 6, 8, [
        ("مانده حساب", f"=$H${11 + N + 2}", "#,##0;[Red]-#,##0", L.TERRA),
        ("جمع گردش", f"=$F${12 + N}+$G${12 + N}", "#,##0;[Red]-#,##0", L.FOREST),
    ])
    ws.merge_cells("A9:B9")
    c = ws["A9"]
    c.value = "کد حساب"
    L.st(c, name=FA, size=10, bold=True, color=L.WHITE, bg=L.FOREST, border=L.FOREST)
    c = ws["C9"]
    c.value = "C001"
    L.st(c, name=EN, size=10, bold=True, color=L.INK, bg=L.INPUT, border=L.TERRA)
    ws.merge_cells("D9:H9")
    c = ws["D9"]
    c.value = ('=IF($C$9="","کد حساب را انتخاب کنید …",'
               'IFERROR(VLOOKUP($C$9,AcctTbl,2,0)&"   ·   "&VLOOKUP($C$9,AcctTbl,3,0),'
               '"⚠ این کد در جدول حساب‌ها نیست"))')
    L.st(c, name=FA, size=9.5, bold=True, color=L.FOREST, bg=L.SAND, border=L.LINE, h="right")
    _dvs(ws, "C9", {"validate": "list", "source": "ListAllCodes",
                    "promptTitle": "کد حساب",
                    "prompt": "کد حسابی که می‌خواهید دفترش را ببینید."})

    L.thead(ws, 11, [("تاریخ", "Date", 12), ("کد سند", "Doc", 12), ("نوع", "Type", 14),
                     ("شرح", "Description", 34), ("گروه", "Group", 18),
                     ("بدهکار", "Debit", 16), ("بستانکار", "Credit", 16),
                     ("مانده", "Balance", 18)], height=30)
    key = f"SMALL(Book!$S$10:$S${B},ROW()-11)"
    pos = f"MATCH({key},Book!$S$10:$S${B},0)"
    for k in range(1, N + 1):
        r = 11 + k
        for c, src in (("A", "A"), ("B", "B"), ("C", "C"), ("D", "E"), ("E", "O"),
                       ("F", "L"), ("G", "M")):
            cell = ws[f"{c}{r}"]
            cell.value = f"=IFERROR(INDEX(Book!${src}$10:${src}${B},{pos}),\"\")"
            L.st(cell, name=EN if c in "ABFG" else FA, size=9, color=L.INK,
                 bg=L.CREAM if k % 2 else L.WHITE,
                 h="right" if c in "DE" else "center",
                 fmt=("#,##0;[Red]-#,##0" if c in "FG" else None))
        c = ws[f"H{r}"]
        c.value = f'=IF($A{r}="","",SUM($F$12:F{r})-SUM($G$12:G{r}))'
        L.st(c, name=EN, size=9, bold=True, color=L.FOREST,
             bg=L.CREAM if k % 2 else L.WHITE, fmt="#,##0;[Red]-#,##0")
        ws.row_dimensions[r].height = 16

    tot, fin = 11 + N + 1, 11 + N + 2
    for row, label in ((tot, "جمع"), (fin, "مانده نهایی")):
        ws.merge_cells(f"A{row}:E{row}")
        c = ws[f"A{row}"]
        c.value = label
        L.st(c, name=FA, size=10, bold=True, color=L.WHITE, bg=L.FOREST,
             border=L.FOREST, h="right")
        f_deb = f"=SUM(F12:F{11 + N})" if row == tot else f"=$F${tot}"
        f_cred = f"=SUM(G12:G{11 + N})" if row == tot else f"=$G${tot}"
        for col, val in (("F", f_deb), ("G", f_cred)):
            c = ws[f"{col}{row}"]
            c.value = val
            L.st(c, name=EN, size=10, bold=True, color=L.FOREST, bg=L.MINT,
                 border=L.LINE, fmt="#,##0;[Red]-#,##0")
        c = ws[f"H{row}"]
        c.value = f"=$F${tot}-$G${tot}"
        L.st(c, name=EN, size=11, bold=True, color=L.TERRA, bg=L.SAND,
             border=L.LINE, fmt="#,##0;[Red]-#,##0")
        ws.row_dimensions[row].height = 22

    L.note(ws, fin + 2, 8,
           "همه ردیف‌های دفتر روزنامه این حساب — از باسکول، سبد، پرداخت، تقسیم مبلغ و کارمزد — "
           "به ترتیب تاریخ. ردیف‌های «تقسیم» و «کارمزد» هم زیر همین حساب می‌آیند تا گردش "
           "واقعی دیده شود.", bg=L.MINT, h=34)
    L.page(ws, 8, fin + 2, landscape=False)
    return "Ledger"
