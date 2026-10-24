# -*- coding: utf-8 -*-
"""Content edits for 01_Entry: the POS gateway control sheet and the
operation rules of the new series (Mostafa / weighbridge income / misc sale).
"""
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

FOREST = "FF14453D"
SAGE = "FF7FA99B"
CREAM = "FFFBF8F1"
SAND = "FFF3EADB"
SAND2 = "FFF7F1E6"
GOLD = "FFC89B3C"
TERRA = "FFC4623F"
MINT = "FFE4F0E8"
RED = "FFB4544A"
GREEN = "FF3F7A57"
MUTED = "FF6B7A76"
LINE = "FFDDD5C4"
WHITE = "FFFFFFFF"

FA = "B Nazanin"
EN = "Tahoma"


def _f(name=FA, size=10, bold=False, color="FF2B2A26", fill=None, border=None,
       align="center", wrap=False, num=None):
    side = Side(style="thin", color=LINE) if border else None
    return dict(
        font=Font(name=name, size=size, bold=bold, color=color),
        fill=PatternFill("solid", fgColor=fill) if fill else None,
        border=Border(left=side, right=side, top=side, bottom=side) if border else None,
        alignment=Alignment(horizontal=align, vertical="center", wrap_text=wrap,
                            readingOrder=2),
        num_format=num,
    )


def style(ws, ref, **kw):
    c = ws[ref]
    st = _f(**kw)
    if st["font"]:
        c.font = st["font"]
    if st["fill"]:
        c.fill = st["fill"]
    if st["border"]:
        c.border = st["border"]
    c.alignment = st["alignment"]
    if st["num_format"]:
        c.number_format = st["num_format"]
    return c


# --------------------------------------------------------------------- rules
def update_ops(wb):
    """Rewrites the four operation rules the client asked for. Everything the
    journal does is driven by this table (Sync!I..W), so no formula changes."""
    ws = wb["Sync"]
    rows = {ws.cell(r, 9).value: r for r in range(10, 34) if ws.cell(r, 9).value}
    edits = {
        # فروش متفرقه: مبلغ به حسابی که کاربر انتخاب میکند (کد حساب = مقصد وجه)
        "فروش متفرقه / صادراتی": {18: "هیچ", 19: "—", 20: None, 21: None},
        # سایر خروج: مبلغ = درآمد باسکول روی ۲۰۰۶، بیارتباط با صاحب بار
        "سایر خروج": {14: "مبلغ", 15: "بدهکار", 16: "X002", 17: "بستانکار",
                      18: "هیچ", 19: "—", 20: None, 21: None, 22: "BN02"},
        # انعام مصطفی: ۲۰۰۶ بدهکار / انعام مصطفی بستانکار
        "انعام مصطفی": {15: "بستانکار", 16: "BN02", 17: "بدهکار", 22: "X001"},
    }
    for title, cols in edits.items():
        r = rows.get(title)
        if not r:
            print("  ! op not found:", title)
            continue
        for col, val in cols.items():
            ws.cell(r, col).value = val
    # keep the active-flag list in sync (column W)
    return sorted(rows.items())


# ----------------------------------------------------------------- POS sheet
def add_pos_sheet(wb, dates):
    ws = wb.create_sheet("POS", 6)          # after Sum, before Sync
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False
    widths = {"A": 13, "B": 13, "C": 16, "D": 17, "E": 17, "F": 15, "G": 14, "H": 30}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    # ---- brand + title
    ws.merge_cells("A1:H1")
    style(ws, "A1", size=10.5, bold=True, color=FOREST, align="right")
    ws["A1"] = "مازندشت  ·  سورتینگ و بسته‌بندی مرکبات"
    ws.merge_cells("A2:H2")
    style(ws, "A2", size=21, bold=True, color=FOREST, align="right")
    ws["A2"] = "کنترل پز باسکول و واریز حساب ۲۰۰۶ صادرات"
    ws.merge_cells("A3:H3")
    style(ws, "A3", size=10.5, color=MUTED, align="right")
    ws["A3"] = ("پز هر روز باسکول  ↔  واریزی روز بعد به ۲۰۰۶  ·  کارت‌به‌کارت به سایر حساب‌ها  ·  "
                "هشدار خودکار ناترازی")
    ws.row_dimensions[1].height = 20
    ws.row_dimensions[2].height = 34
    ws.row_dimensions[3].height = 20

    ws.merge_cells("A4:H4")
    style(ws, "A4", size=9.5, color="FF3A3327", fill=SAND, align="right", wrap=True, border=True)
    ws["A4"] = ("قاعده: هر مبلغی که از باسکول به حساب ۲۰۰۶ صادرات (پوز) می‌رود، همان روز در ستون «پز باسکول» جمع می‌شود. "
                "روز بعد که پول به حساب شرکت واریز یا کارت‌به‌کارت می‌شود، در شیت Pay با نوع «دریافت» ثبت کنید تا اینجا خودکار بیاید. "
                "اگر اختلاف غیرصفر شد، یعنی یک واریزی ثبت نشده یا یک پز اشتباه وارد شده است.")
    ws.row_dimensions[4].height = 42

    # ---- KPI cards
    cards = [
        ("جمع پز باسکول", "=SUM($C$12:$C$51)", FOREST),
        ("جمع واریز به ۲۰۰۶", "=SUM($D$12:$D$51)", GREEN),
        ("کارت‌به‌کارت / سایر", "=SUM($E$12:$E$51)", GOLD),
        ("اختلاف کل", "=$A$7-$C$7-$E$7", TERRA),
    ]
    spans = [("A", "B"), ("C", "D"), ("E", "F"), ("G", "H")]
    for (lab, formula, color), (c0, c1) in zip(cards, spans):
        ws.merge_cells(f"{c0}6:{c1}6")
        style(ws, f"{c0}6", size=8, bold=True, color=MUTED, fill=SAND2, border=True)
        ws[f"{c0}6"] = lab
        ws.merge_cells(f"{c0}7:{c1}7")
        style(ws, f"{c0}7", size=18, bold=True, color=color, fill=CREAM, border=True, num="#,##0")
        ws[f"{c0}7"] = formula
    ws.row_dimensions[6].height = 20
    ws.row_dimensions[7].height = 30

    # ---- section
    ws.merge_cells("A9:H9")
    style(ws, "A9", size=12, bold=True, color=WHITE, fill=FOREST, align="right")
    ws["A9"] = "۰۱   روز‌های پز — پز باسکول در برابر واریزی روز بعد"

    heads = [
        ("تاریخ پز", "Pay date"), ("تاریخ واریز", "Deposit date"),
        ("پز باسکول (ریال)", "POS in scale"), ("واریز به ۲۰۰۶", "To 2006"),
        ("کارت‌به‌کارت", "Card to card"), ("اختلاف", "Difference"),
        ("وضعیت", "Status"), ("توضیح", "Note"),
    ]
    for i, (fa, en) in enumerate(heads, start=1):
        c = ws.cell(10, i, fa)
        style(ws, c.coordinate, size=9.5, bold=True, color=WHITE, fill=FOREST, border=True, wrap=True)
        c2 = ws.cell(11, i, en)
        style(ws, c2.coordinate, name=EN, size=7, color="FFD8E6DF", fill=FOREST, border=True)
    ws.row_dimensions[10].height = 32
    ws.row_dimensions[11].height = 13

    for i, d in enumerate(dates[:40]):
        r = 12 + i
        ws.row_dimensions[r].height = 20
        style(ws, f"A{r}", name=EN, size=9, fill="FFFFF4D6", border=True)
        ws[f"A{r}"] = d
        style(ws, f"B{r}", name=EN, size=9, fill="FFFFF4D6", border=True)
        style(ws, f"C{r}", name=EN, size=9, fill="FFF6F5F0", border=True, num="#,##0", color=MUTED)
        ws[f"C{r}"] = (f'=IF($A{r}="","",SUMIFS(Book!$L$10:$L$1809,Book!$A$10:$A$1809,$A{r},'
                       f'Book!$D$10:$D$1809,"BN02",Book!$C$10:$C$1809,"باسکول"))')
        style(ws, f"D{r}", name=EN, size=9, fill="FFF6F5F0", border=True, num="#,##0", color=MUTED)
        ws[f"D{r}"] = (f'=IF($A{r}="","",SUMIFS(Pay!$K$10:$K$159,Pay!$B$10:$B$159,'
                       f'IF($B{r}="",$A{r},$B{r}),Pay!$D$10:$D$159,"دریافت",Pay!$E$10:$E$159,"کارت/پوز",Pay!$I$10:$I$159,"BN02"))')
        style(ws, f"E{r}", name=EN, size=9, fill="FFF6F5F0", border=True, num="#,##0", color=MUTED)
        ws[f"E{r}"] = (f'=IF($A{r}="","",SUMIFS(Pay!$K$10:$K$159,Pay!$B$10:$B$159,'
                       f'IF($B{r}="",$A{r},$B{r}),Pay!$D$10:$D$159,"دریافت",Pay!$E$10:$E$159,"کارت/پوز",Pay!$I$10:$I$159,"<>BN02"))')
        style(ws, f"F{r}", name=EN, size=9, bold=True, fill=CREAM, border=True, num="#,##0;[Red]-#,##0")
        ws[f"F{r}"] = f'=IF($A{r}="","",$C{r}-$D{r}-$E{r})'
        style(ws, f"G{r}", size=9.5, bold=True, fill=CREAM, border=True)
        ws[f"G{r}"] = f'=IF($A{r}="","",IF(ABS($F{r})<=1000,"✓ تراز","⚠ بررسی شود"))'
        style(ws, f"H{r}", size=9, fill="FFFFF4D6", border=True, align="right")

    # totals
    ws.merge_cells("A52:B52")
    style(ws, "A52", size=10, bold=True, color=WHITE, fill=GOLD, border=True)
    ws["A52"] = "جمع کل"
    for col in ("C", "D", "E", "F"):
        style(ws, f"{col}52", name=EN, size=10, bold=True, color=WHITE, fill=GOLD,
              border=True, num="#,##0;[Red]-#,##0")
        ws[f"{col}52"] = f"=SUM({col}12:{col}51)"
    style(ws, "G52", size=9.5, bold=True, color=WHITE, fill=GOLD, border=True)
    ws["G52"] = '=IF(ABS($F$52)<=1000,"✓ تراز","⚠ ناتراز")'
    ws.row_dimensions[52].height = 24

    # ---- enaam block
    ws.merge_cells("A54:H54")
    style(ws, "A54", size=12, bold=True, color=WHITE, fill=FOREST, align="right")
    ws["A54"] = "۰۲   کنترل انعام مصطفی — ۲۰۰۶ صادرات در برابر حساب انعام مصطفی"
    rows = [
        ("انعام ثبت‌شده در باسکول (بستانکار مصطفی)",
         '=SUMIFS(Book!$M$10:$M$1809,Book!$D$10:$D$1809,"X001",Book!$C$10:$C$1809,"باسکول")'),
        ("پرداخت‌ها/انتقال‌های بانکی به مصطفی",
         '=SUMIFS(Pay!$K$10:$K$159,Pay!$F$10:$F$159,"X001",Pay!$D$10:$D$159,"پرداخت")'),
        ("مانده حساب انعام مصطفی (بستانکار − بدهکار)",
         '=SUMIFS(Book!$M$10:$M$1809,Book!$D$10:$D$1809,"X001")-SUMIFS(Book!$L$10:$L$1809,Book!$D$10:$D$1809,"X001")'),
        ("کنترل ۲۰۰۶ ↔ مصطفی (باید صفر باشد)",
         '=SUMIFS(Book!$L$10:$L$1809,Book!$D$10:$D$1809,"BN02",Book!$N$10:$N$1809,"انعام مصطفی")'
         '-SUMIFS(Book!$M$10:$M$1809,Book!$D$10:$D$1809,"X001",Book!$N$10:$N$1809,"انعام مصطفی")'),
    ]
    for i, (lab, formula) in enumerate(rows):
        r = 56 + i
        ws.merge_cells(f"A{r}:D{r}")
        style(ws, f"A{r}", size=10, bold=(i == 3), color="FF2B2A26", fill=SAND2 if i % 2 == 0 else WHITE,
              border=True, align="right")
        ws[f"A{r}"] = lab
        ws.merge_cells(f"E{r}:F{r}")
        style(ws, f"E{r}", name=EN, size=10, bold=True, color=FOREST, fill=CREAM, border=True,
              num="#,##0;[Red]-#,##0")
        ws[f"E{r}"] = formula
        style(ws, f"G{r}", size=9.5, bold=True, border=True, fill=CREAM)
        ws[f"G{r}"] = f'=IF(ROUND($E{r},0)=0,"✓","⚠")'
        ws.row_dimensions[r].height = 22

    # ---- guide
    ws.merge_cells("A62:H62")
    style(ws, "A62", size=12, bold=True, color=WHITE, fill=SAGE, align="right")
    ws["A62"] = "۰۳   راهنما — این شیت را چطور بخوانم؟"
    guide = [
        "۱) ستون «پز باسکول» خودکار است: هر خط دفتر (شیت Book) که از باسکول به حساب BN02 (صادرات ۲۰۰۶) بدهکار شده باشد را جمع می‌کند.",
        "۲) واریزی روز بعد را در شیت Pay ثبت کنید: نوع = دریافت، روش = بانک، بانک/مقصد = BN02 (صادرات ۲۰۰۶).",
        "۳) اگر پول به جای ۲۰۰۶ به حساب بانک دیگری کارت‌به‌کارت شده، همان را در Pay با بانک/مقصد = آن بانک ثبت کنید؛ در ستون کارت‌به‌کارت می‌آید.",
        "۴) اگر روز واریز با روز پز یکی نیست، تاریخ واریز را در ستون B بنویسید؛ جمع‌ها بر اساس همان تاریخ محاسبه می‌شوند.",
        "۵) ستون وضعیت با اختلاف بیش از ۱٬۰۰۰ ریال هشدار می‌دهد؛ برای دیدن ریز، شیت Book را با فیلتر BN02 و منبع باسکول ببینید.",
        "۶) انعام مصطفی: هر ردیف باسکول با نوع «انعام مصطفی» خودکار ۲۰۰۶ را بدهکار و حساب X001 (انعام مصطفی) را بستانکار می‌کند؛ ردیف کنترل باید صفر بماند.",
    ]
    for i, txt in enumerate(guide):
        r = 63 + i
        ws.merge_cells(f"A{r}:H{r}")
        style(ws, f"A{r}", size=9.5, color="FF2B2A26", fill=CREAM if i % 2 == 0 else WHITE, border=True,
              align="right", wrap=True)
        ws[f"A{r}"] = txt
        ws.row_dimensions[r].height = 24

    # ---- print
    ws.print_area = "A1:H69"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "portrait"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.oddHeader.left.text = "مازندشت — سامانه حسابداری سورتینگ"
    ws.oddHeader.center.text = "MAZANDASHT · POS CONTROL"
    ws.oddHeader.right.text = "&A"
    ws.oddFooter.left.text = "&F"
    ws.oddFooter.center.text = "صفحه &P از &N"
    ws.oddFooter.right.text = "Excel 2013"
    return ws


# -------------------------------------------------------------------- labels
def update_texts(wb):
    subs = [
        ("01_Register.xlsx", "01_Entry.xlsx"),
        ("INFO.xlsx", "00_Master.xlsx"),
        ("04_Invoice.xlsx", "04_Inv.xlsx"),
        ("05_Balance.xlsx", "05_Bal.xlsx"),
        ("06_LoadBuyers.xlsx", "06_Buyers.xlsx"),
        ("07_Sample.xlsx", "07_Demo.xlsx"),
    ]
    n = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and any(a in c.value for a in
                                                    ("01_Register", "INFO.xlsx", "04_Invoice",
                                                     "05_Balance", "06_LoadBuyers", "07_Sample")):
                    v = c.value
                    for a, b in subs:
                        v = v.replace(a, b)
                    c.value = v
                    n += 1
    # scale headers: the account column becomes the money destination for misc sales
    ws = wb["Baskol"]
    if ws.cell(8, 7).value:
        ws.cell(8, 7).value = "کد حساب / مقصد وجه"
        ws.cell(9, 7).value = "Account / Money-to"
    if ws.cell(8, 17).value:
        ws.cell(8, 17).value = "مبلغ پوز/باسکول (درآمد)"
        ws.cell(9, 17).value = "POS / Fee amount"
    return n


def home_strip(wb):
    ws = wb["Home"]
    r = ws.max_row + 2
    ws.merge_cells(f"A{r}:J{r}")
    style(ws, f"A{r}", size=10.5, color="FF2A3B35", fill=MINT, border=True, align="right", wrap=True)
    ws.cell(r, 1).value = ("شیت POS  ·  کنترل روزانه پز باسکول در برابر واریزی روز بعد به حساب ۲۰۰۶ صادرات؛ "
                           "اگر تراز نباشد، هشدار ⚠ می‌دهد.")
    ws.row_dimensions[r].height = 26
    r2 = r + 1
    ws.merge_cells(f"A{r2}:J{r2}")
    style(ws, f"A{r2}", size=10.5, color="FF2A3B35", fill=SAND, border=True, align="right", wrap=True)
    ws.cell(r2, 1).value = ("قاعده جدید: هر مبلغی که برای خروج بار ثبت شود درآمد باسکول است و به ۲۰۰۶ می‌رود؛ "
                            "برای فروش متفرقه، ستون «کد حساب / مقصد وجه» همان حسابی است که پول به آن واریز شده.")
    ws.row_dimensions[r2].height = 30
    ws.print_area = f"A1:J{r2}"


# ------------------------------------------------------------------ help text
HELP_BLOCK = [
    ("۰۹", "قواعد حسابداری نسخه ۲ (این نسخه)"),
    ("p", "این بخش قواعدی را توضیح می‌دهد که در نسخه جدید اضافه شده‌اند. همه این قواعد با همان جدول عملیات شیت Sync کار می‌کنند؛ "
          "یعنی اگر روزی نوع عملیات جدیدی اضافه کردید، کافی است همان چهار ستون بدهکار/بستانکار را پر کنید."),
    ("b", "انعام مصطفی: حساب «انعام مصطفی» یک حساب مستقل با کد X001 است. وقتی در باسکول نوع عملیات «انعام مصطفی» را انتخاب کنید، "
          "سیستم خودکار حساب صادرات ۲۰۰۶ (BN02) را بدهکار و حساب انعام مصطفی را بستانکار می‌کند؛ پس مانده ۲۰۰۶ و مصطفی همیشه "
          "به اندازه هم جابه‌جا می‌شود و ردیف کنترل در شیت POS صفر می‌ماند."),
    ("b", "مبلغ خروجی بار: هر مبلغی که برای خروج بار (به‌جز فروش به بارفروش که فاکتور واقعی دارد) ثبت شود، درآمد باسکول است و "
          "به حساب ۲۰۰۶ صادرات می‌رود؛ بدهی یا طلب صاحب بار هیچ تغییری نمی‌کند. برای همین در نوع عملیات «سایر خروج»، حساب پیش‌فرض "
          "روی BN02 تنظیم شده و طرف مقابل آن X002 (درآمد باسکول) است."),
    ("b", "فروش متفرقه: در ستون «کد حساب / مقصد وجه»، حسابی را انتخاب کنید که پول به آن واریز شده است (هر بانک یا صندوق). "
          "طرف مقابل این فروش، حساب X006 (فروش مرکبات) است. اگر خریدار با پوز ۲۰۰۶ پرداخت کرده، همان BN02 را انتخاب کنید."),
    ("b", "شیت POS: کنترل روزانه پز باسکول در برابر واریزی روز بعد است. واریزی را در شیت Pay با نوع «دریافت»، روش «کارت/پوز» و "
          "بانک مقصد ثبت کنید. اگر پول به بانک دیگری کارت‌به‌کارت شده، همان بانک را انتخاب کنید تا در ستون کارت‌به‌کارت بیاید. "
          "اختلاف بیش از ۱٬۰۰۰ ریال با ⚠ نشان داده می‌شود."),
    ("b", "چیدمان: همه شیت‌ها راست‌به‌چپ، اندازه کاغذ A4، سربرگ و فوتر چاپی دارند. نام شیت‌ها و فایل‌ها کوتاه و انگلیسی است."),
]


def add_help(wb, note_style_src):
    ws = wb["Help"]
    last = ws.max_row + 2
    for i, (kind, text) in enumerate(HELP_BLOCK):
        r = last + i
        if kind != "p":
            ws.merge_cells(f"A{r}:J{r}")
            c = ws.cell(r, 1)
            c.value = ("۰۹  " if kind == "۰۹" else "•  ") + (text if kind != "۰۹" else text)
            style(ws, f"A{r}", size=10.5 if kind == "b" else 11, bold=(kind == "۰۹"),
                  color=WHITE if kind == "۰۹" else "FF2B2A26",
                  fill=FOREST if kind == "۰۹" else MINT, border=True, align="right", wrap=True)
            ws.row_dimensions[r].height = 44 if kind == "b" else 26
        else:
            ws.merge_cells(f"A{r}:J{r}")
            style(ws, f"A{r}", size=10.5, color="FF2B2A26", border=False, align="right", wrap=True)
            ws.cell(r, 1).value = text
            ws.row_dimensions[r].height = 40
    ws.print_area = f"A1:J{last + len(HELP_BLOCK) - 1}"
    return last + len(HELP_BLOCK)


def build(src, dst, dates):
    wb = openpyxl.load_workbook(src)
    update_ops(wb)
    n = update_texts(wb)
    add_pos_sheet(wb, dates)
    add_help(wb, None)
    home_strip(wb)
    wb.calculation.fullCalcOnLoad = True
    wb.active = 0
    wb.save(dst)
    print(f"  01_Entry: {n} text cells updated, POS sheet added, OPS rules updated")
    return dst
