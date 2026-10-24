# -*- coding: utf-8 -*-
"""Post-processing helpers for the v03 files (code unification + journal fixes).

Everything the v01-derived workbooks carry as a *text* code is normalised here,
so the whole set speaks one language:

    BANK01 … BANK10   →   BN01 … BN10        (کد بانک = کد حساب معین)

The journal (01_Entry) additionally gets:

* every bank leg routed through the clearing account **X011 «واسط بانک»**, so
  the bank accounts are only ever posted by 02_Bank — no movement counted twice;
* a hidden «بانک مرجع» column (P) that remembers which bank the money went
  through, used to build the bank statement of 02_Bank;
* guards that drop empty / zero rows out of the journal instead of writing `0`
  into the ledger.
"""
import re
from copy import copy

import openpyxl
from openpyxl.styles import Alignment, Font
from openpyxl.workbook.defined_name import DefinedName

BANK_RE = re.compile(r"BANK(\d{2})")
CLEARING = "X011"
POS_BANK = "BN02"
BANK_LINK = "[02_Bank.xlsx]"
BANK_CODES = {"BN%02d" % i for i in range(1, 11)} | {"BANK%02d" % i for i in range(1, 11)}


# --------------------------------------------------------------- code unify
def unify_codes(path, save=True):
    """BANKxx → BNxx in values, formulas, defined names and validations."""
    wb = openpyxl.load_workbook(path)
    hits = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str) and "BANK" in v:
                    new = BANK_RE.sub(r"BN\1", v)
                    if new != v:
                        cell.value = new
                        hits += 1

    for name in list(wb.defined_names.keys()):
        dn = wb.defined_names[name]
        txt = getattr(dn, "attr_text", None)
        if txt and "BANK" in txt:
            fixed = BANK_RE.sub(r"BN\1", txt)
            if fixed != txt:
                del wb.defined_names[name]
                wb.defined_names.add(DefinedName(name, attr_text=fixed))
                hits += 1

    for ws in wb.worksheets:
        for dv in getattr(ws.data_validations, "dataValidation", []):
            for attr in ("formula1", "formula2"):
                v = getattr(dv, attr, None)
                if isinstance(v, str) and "BANK" in v:
                    setattr(dv, attr, BANK_RE.sub(r"BN\1", v))
                    hits += 1

    if save:
        wb.calculation.fullCalcOnLoad = True
        wb.save(path)
    return hits


# ------------------------------------------------------------ entry patches
REF_RE = re.compile(r"^=IF\((\w+)!\$([A-Z]{1,2})(\d+)=\"\",\"\",")


def entry_clearing(path, save=True):
    """Route every bank leg of the journal through X021 and tag the bank."""
    wb = openpyxl.load_workbook(path)
    sy = wb["Sync"]
    changed = 0

    # --- 1) operation rules: P ledger · T pos-bank · V fixed party
    for r in range(10, 34):
        title = sy.cell(r, 9).value
        if not isinstance(title, str) or not title.strip():
            continue
        for col in (16, 20, 22):
            v = sy.cell(r, col).value
            if isinstance(v, str) and v.strip() in BANK_CODES:
                sy.cell(r, col).value = CLEARING
                changed += 1

    # «انعام مصطفی» = his share is *paid out of* the 2006 pot, so the pot is
    # credited and his own account is debited (he received the money).  The
    # statement line then settles the clearing against the bank.
    for r in range(10, 34):
        if str(sy.cell(r, 9).value).strip() == "انعام مصطفی":
            sy.cell(r, 16).value = "X001"        # P ledger  → his account
            sy.cell(r, 17).value = "بدهکار"      # Q side    → debit
            sy.cell(r, 22).value = CLEARING      # V party   → clearing
            sy.cell(r, 15).value = "بستانکار"    # O side
            changed += 4

    # «سایر خروج» = money that lands in 2006 → credit X002, party = clearing
    for r in range(10, 34):
        if str(sy.cell(r, 9).value).strip() == "سایر خروج":
            sy.cell(r, 15).value = "بدهکار"       # O party side
            sy.cell(r, 16).value = "X002"         # P ledger
            sy.cell(r, 17).value = "بستانکار"     # Q ledger side
            sy.cell(r, 22).value = CLEARING       # V fixed party
            changed += 4

    # --- 2) the bank dropdown must offer codes (the ledger key)
    if "ListBanks" in wb.defined_names:
        del wb.defined_names["ListBanks"]
    wb.defined_names.add(DefinedName("ListBanks", attr_text="Sync!$AK$10:$AK$19"))

    # --- 2b) an internal transfer between our own banks is a bank-only event:
    #         the journal keeps the instruction but posts nothing (02_Bank books it)
    pay0 = wb["Pay"]
    for r in range(10, 160):
        if str(pay0.cell(r, 5).value).strip() == "انتقال داخلی":
            for col in (30, 32):                     # AD, AF
                f = pay0.cell(r, col).value
                if isinstance(f, str) and f.startswith("=IF("):
                    pay0.cell(r, col).value = '=IF($E%d="انتقال داخلی","","")' % r
                    changed += 1

    # --- 3) Pay: every bank leg is booked through the clearing account
    #     AD = first line (debit), AF = second line (credit)
    pay = wb["Pay"]
    for r in range(10, 160):
        for col in (28, 30, 32):                     # AB, AD, AF
            f = pay.cell(r, col).value
            if not isinstance(f, str) or "BankTbl" not in f:
                continue
            new = f
            for key in ("I", "N"):
                pat = r"IFERROR\(VLOOKUP\(\$%s%d,BankTbl,4,0\),\$%s%d\)" % (key, r, key, r)
                new = re.sub(pat, '"%s"' % CLEARING, new)
            if new != f:
                pay.cell(r, col).value = new
                changed += 1

    # --- 4) Book: hidden «بانک مرجع» column P
    bk = wb["Book"]
    bk.cell(8, 16, "بانک مرجع")
    bk.cell(9, 16, "Bank ref.")
    bk.column_dimensions["P"].width = 11
    for r in range(10, 1810):
        a = bk.cell(r, 1).value
        if not isinstance(a, str):
            continue
        m = REF_RE.match(a)
        if not m:
            continue
        sheet, col, row = m.group(1), m.group(2), int(m.group(3))
        if sheet == "Scale":
            bk.cell(r, 16, f'=IF(Scale!${col}{row}<>"{CLEARING}","","{POS_BANK}")')
        elif sheet == "Pay":
            bk.cell(r, 16,
                    f'=IF(Pay!${col}{row}<>"{CLEARING}","",IF(Pay!$I{row}="","",Pay!$I{row}))')
        else:
            bk.cell(r, 16, "")

    if save:
        wb.calculation.fullCalcOnLoad = True
        wb.save(path)
    return changed


def entry_zero_guard(path, save=True):
    """Never write a journal row when its key account or amount is empty/zero.

    Every cell of a Book row is `=IF(<source>="","",<value>)`; a scale row whose
    account resolves to `0` used to leave a ghost line in the ledger.  The guard
    adds `<source>=0` to the test, so the whole row stays blank.
    """
    wb = openpyxl.load_workbook(path)
    bk = wb["Book"]
    changed = 0
    pat = re.compile(r"^=IF\((\w+!\$[A-Z]+\d+)=\"\",\"\",(.+)\)$")
    for r in range(10, 1810):
        for col in range(1, 16):
            cell = bk.cell(r, col)
            f = cell.value
            if not isinstance(f, str):
                continue
            m = pat.match(f)
            if not m:
                continue
            key, val = m.group(1), m.group(2)
            if "OR(%s=" % key in f:
                continue
            cell.value = f'=IF(OR({key}="",{key}=0),"",{val})'
            changed += 1
    if save:
        wb.calculation.fullCalcOnLoad = True
        wb.save(path)
    return changed


# ------------------------------------------------------------------ 05_Bal
def entry_pos(path, save=True):
    """Rewrite the POS sheet of 01_Entry as a *POS register*.

    In v03 the journal posts POS takings to the clearing account X011 and the
    bank statement lives in 02_Bank, so a "POS ↔ next-day deposit" check cannot
    be computed inside this file any more (it would flag every single day).
    The sheet becomes a clean daily register; the two-sided reconciliation is
    in 02_Bank!Card and in the Pos sheet of 08_Dash.
    """
    wb = openpyxl.load_workbook(path)
    ws = wb["POS"]
    C = CLEARING
    BOOK_A = "Book!$A$10:$A$1809"
    BOOK_C = "Book!$C$10:$C$1809"
    BOOK_D = "Book!$D$10:$D$1809"
    BOOK_L = "Book!$L$10:$L$1809"
    BOOK_M = "Book!$M$10:$M$1809"
    BOOK_N = "Book!$N$10:$N$1809"

    ws["A2"].value = "پز باسکول — دفتر روزانه"
    ws["A3"].value = ("پز هر روز باسکول  ·  تجمعی و سهم  ·  تطبیق با واریز بانکی در فایل "
                      "02_Bank (شیت Card) و داشبورد (شیت Pos)")
    ws["A4"].value = ("قاعده: پز باسکول در دفتر روزنامه به حساب واسط X011 (در راه) می‌رود و واریز آن "
                      "در فایل 02_Bank ثبت می‌شود؛ هیچ سندی مستقیم روی حساب بانک (BNxx) زده نمی‌شود. "
                      "تطبیق دو طرف را در 02_Bank!Card ببینید.")
    ws["A6"].value, ws["A7"].value = "جمع پز باسکول", "=SUM($B$12:$B$51)"
    ws["C6"].value, ws["C7"].value = "روزهای دارای پز", '=COUNTIF($B$12:$B$51,">0")'
    ws["E6"].value, ws["E7"].value = "میانگین روزانه", "=IFERROR($A$7/$C$7,0)"
    ws["G6"].value, ws["G7"].value = "بیشترین روز", "=MAX($B$12:$B$51)"

    ws["A9"].value = "۰۱   دفتر روزانه پز باسکول"
    heads = [("A", "تاریخ", "Date"), ("B", "پز باسکول (ریال)", "POS"),
             ("C", "تجمعی پز", "Cumulative"), ("D", "سهم از کل", "Share"),
             ("E", "تعداد سند", "Docs"), ("F", "میانگین سند", "Avg"),
             ("G", "وضعیت", "Status"), ("H", "توضیح", "Note")]
    for col, fa, en in heads:
        ws[f"{col}10"].value = fa
        ws[f"{col}11"].value = en

    for r in range(12, 52):
        ws[f"B{r}"].value = (f'=IF($A{r}="","",SUMIFS({BOOK_L},{BOOK_A},$A{r},'
                             f'{BOOK_D},"{C}",{BOOK_N},"پوز"))')
        ws[f"C{r}"].value = f'=IF($A{r}="","",SUM($B$12:$B{r}))'
        ws[f"D{r}"].value = (f'=IF($A{r}="","",IFERROR($B{r}/SUM($B$12:$B$51),0))')
        ws[f"E{r}"].value = (f'=IF($A{r}="","",COUNTIFS({BOOK_A},$A{r},{BOOK_D},"{C}",'
                             f'{BOOK_N},"پوز"))')
        ws[f"F{r}"].value = f'=IF($A{r}="","",IFERROR($B{r}/$E{r},0))'
        ws[f"G{r}"].value = f'=IF($A{r}="","",IF($B{r}>0,"✓ ثبت شد","—"))'
        ws[f"D{r}"].number_format = "0.0%"
        ws[f"F{r}"].number_format = "#,##0"

    # A52:B52 is merged for the «جمع کل» label — free column B for the total
    for rng in [str(m) for m in ws.merged_cells.ranges]:
        if rng == "A52:B52":
            ws.unmerge_cells(rng)
    a52 = ws["A52"]
    b52 = ws["B52"]
    b52.font = copy(a52.font)
    b52.fill = copy(a52.fill)
    b52.border = copy(a52.border)
    b52.alignment = copy(a52.alignment)
    ws["B52"].value = "=SUM(B12:B51)"
    ws["C52"].value = "=$B$52"
    ws["D52"].value = "=SUM(D12:D51)"
    ws["E52"].value = "=SUM(E12:E51)"
    ws["F52"].value = "=IFERROR($B$52/$E$52,0)"
    ws["G52"].value = '=IF($B$52>0,"✓","—")'
    ws["D52"].number_format = "0.0%"
    ws["F52"].number_format = "#,##0"

    # «انعام مصطفی» — everything paid to him must leave through the clearing
    ws["A54"].value = "۰۲   کنترل انعام مصطفی — پرداخت از واسط در برابر حساب X001"
    ws["A57"].value = "انعام پرداخت‌شده از باسکول (بدهکار حساب X001)"
    ws["E57"].value = f'=SUMIFS({BOOK_L},{BOOK_D},"X001",{BOOK_C},"باسکول")'
    ws["A59"].value = "کنترل: پرداخت به مصطفی = بستانکار واسط با نوع «انعام مصطفی»"
    ws["E59"].value = f'=SUMIFS({BOOK_M},{BOOK_D},"{C}",{BOOK_N},"انعام مصطفی")'
    ws["G59"].value = '=IF(ROUND($E$57-$E$59,0)=0,"✓","⚠")'

    notes = [
        "۱) ستون «پز باسکول» خودکار است: هر خط دفتر (شیت Book) که در باسکول با نوع «پوز» به حساب واسط X011 بدهکار شده باشد.",
        "۲) واریز پز به بانک را در فایل 02_Bank با نوع «واریز پز باسکول (۲۰۰۶)» ثبت کنید؛ همان‌جا بانک بدهکار و واسط بستانکار می‌شود.",
        "۳) تطبیق پز با واریزی بانکی در دو جا دیده می‌شود: شیت Card فایل 02_Bank (به تفکیک بانک) و شیت Pos داشبورد 08 (روز به روز).",
        "۴) مانده حساب واسط X011 در پایان ماه باید صفر شود؛ کنترل آن در شیت Audit فایل 05_Bal انجام می‌شود (دفتر + بانک).",
        "۵) اگر روز واریز با روز پز یکی نیست، تاریخ واریز را در فایل بانک بزنید؛ این شیت بر اساس تاریخ پز جمع می‌زند.",
        "۶) انعام مصطفی: هر ردیف باسکول با نوع «انعام مصطفی» واسط X011 را بستانکار و حساب X001 را بدهکار می‌کند؛ ردیف کنترل ۰۲ آن را می‌سنجد.",
    ]
    for i, text in enumerate(notes):
        ws[f"A{63 + i}"].value = text

    # Help sheet — the journal no longer posts to a bank account
    hp = wb["Help"]
    swaps = [
        ("پوز باسکول به حساب صادرات ۲۰۰۶ (BN02) وصل است.",
         "پز باسکول به حساب واسط X011 (در راه) وصل است و از آن‌جا در فایل 02_Bank به حساب صادرات ۲۰۰۶ (BN02) می‌رسد."),
        ("بدهکار صادرات ۲۰۰۶ / بستانکار درآمد باسکول",
         "بدهکار واسط X011 / بستانکار درآمد باسکول"),
        ("بدهکار صادرات ۲۰۰۶ / بستانکار انعام مصطفی",
         "بدهکار واسط X011 / بستانکار انعام مصطفی"),
        ("سیستم خودکار حساب صادرات ۲۰۰۶ (BN02) را بدهکار",
         "سیستم خودکار حساب واسط X011 را بدهکار"),
        ("حساب پیش‌فرض روی BN02 تنظیم شده", "حساب پیش‌فرض روی X011 تنظیم شده"),
        ("اگر خریدار با پوز ۲۰۰۶ پرداخت کرده، همان BN02 را انتخاب کنید.",
         "اگر خریدار با پز ۲۰۰۶ پرداخت کرده، همان X011 را انتخاب کنید؛ واریز بانکی‌اش در 02_Bank ثبت می‌شود."),
        ("به حساب ۲۰۰۶ صادرات می‌رود", "به حساب واسط X011 و از آن‌جا به ۲۰۰۶ صادرات می‌رود"),
    ]
    for row in hp.iter_rows():
        for cell in row:
            if not isinstance(cell.value, str):
                continue
            text = cell.value
            for old, new in swaps:
                if old in text:
                    text = text.replace(old, new)
            if text != cell.value:
                cell.value = text

    if save:
        wb.calculation.fullCalcOnLoad = True
        wb.save(path)
    return 1


def bal_ibank(path, save=True):
    """IBank of the balance sheet reads the bank sub-ledger (02_Bank), not the journal.

    Since the journal posts bank legs to the clearing account only, the real
    balance of every BNxx account lives in 02_Bank!Post.  IBank is repointed
    there (relative external link, same folder) and labelled accordingly.
    """
    wb = openpyxl.load_workbook(path)
    ws = wb["IBank"]
    n = 0
    for i, r in enumerate(range(12, 22), start=1):
        code = ws.cell(r, 2).value
        if not code:
            continue
        # name / code / opening always come from the mother data (BankTbl)
        ws.cell(r, 1).value = f"=INDEX(BankTbl,{i},2)"
        ws.cell(r, 2).value = f"=INDEX(BankTbl,{i},1)"
        ws.cell(r, 3).value = f"=INDEX(BankTbl,{i},3)"
        ws.cell(r, 4).value = f"=SUMIF('{BANK_LINK}Post'!$D$13:$D$292,$B{r},'{BANK_LINK}Post'!$I$13:$I$292)"
        ws.cell(r, 5).value = f"=SUMIF('{BANK_LINK}Post'!$D$13:$D$292,$B{r},'{BANK_LINK}Post'!$J$13:$J$292)"
        n += 1
    for row in ws.iter_rows(min_row=1, max_row=4):
        for cell in row:
            if isinstance(cell.value, str) and "از دفتر" in cell.value:
                cell.value = cell.value.replace(
                    "از دفتر", "از دفتر تراکنش‌های بانکی (02_Bank)")

    # the clearing account closes only when journal + bank ledger are read
    # together — audit it here instead of inside the journal alone
    au = wb["Audit"]
    au["A15"].value = "تسویه کامل واسط X011 (دفتر روزنامه + دفتر بانک)"
    p_ = f"'{BANK_LINK}Post'"
    au["E15"].value = (
        '=IF(ROUND((SUMIF(Data!$C$10:$C$1807,"X011",Data!$G$10:$G$1807)'
        '-SUMIF(Data!$C$10:$C$1807,"X011",Data!$H$10:$H$1807))'
        f'+(SUMIF({p_}!$D$13:$D$292,"X011",{p_}!$I$13:$I$292)'
        f'-SUMIF({p_}!$D$13:$D$292,"X011",{p_}!$J$13:$J$292)),0)=0,"✓","⚠")')

    # Help: bank balances come from the bank book, not from the journal
    hb = wb["Help"]
    for row in hb.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and "برای تطبیق با صورت‌حساب از فایل بانک" in cell.value:
                cell.value = ("مانده هر بانک از دفتر تراکنش‌های بانکی (02_Bank، شیت Post) خوانده می‌شود: "
                              "افتتاحیه از فایل مادر و گردش از همان دفتر. جمع آن با شیت Bank فایل بانک "
                              "و داشبورد (08) یکی است. شیت B4 و B6 فقط گردش دفتر روزنامه را نشان می‌دهند.")

    # the overview card must show the real bank balance (opening + movement)
    ov = wb["Ov"]
    for row in ov.iter_rows(min_row=1, max_row=14):
        for cell in row:
            if isinstance(cell.value, str) and cell.value.strip() == "بانک و صندوق":
                cell.value = "مانده بانک‌ها (02_Bank)"
                ov.cell(cell.row + 1, cell.column).value = "=IBank!$F$22"
            elif (isinstance(cell.value, str) and "بانک" in cell.value
                  and "Data!$E$10" in cell.value):
                cell.value = "=IBank!$F$22"

    # the group trial balances are period *movement*; say so, next to them
    note = ("ارقام این شیت گردش دوره است (بدون مانده اولیه). مانده واقعی هر بانک "
            "در شیت IBank دیده می‌شود و جمع آن با جمع بانک‌های 02_Bank یکی است.")
    grey = Font(name="B Nazanin", size=8, italic=True, color="7A7A7A")
    for sheet, row, last in (("B4", 24, "E"), ("B6", 18, "F")):
        wsx = wb[sheet]
        wsx.merge_cells(f"A{row}:{last}{row}")
        c = wsx.cell(row, 1, note)
        c.font = grey
        c.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
        wsx.row_dimensions[row].height = 22
    if save:
        wb.calculation.fullCalcOnLoad = True
        wb.save(path)
    return n
