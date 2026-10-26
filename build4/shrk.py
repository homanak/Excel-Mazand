# -*- coding: utf-8 -*-
"""v04.5 — شراکت (partnership).

What this module adds on top of v04.4:

1. ``00_Master``  → sheet **Shrk**: the partner register (code SHxx, name,
   «نسبت سهم شریک ٪», «نسبت سهم شرکت ٪» = 100 − it, subject, phones, bank,
   status, note).  The same 20 codes are appended to ``Acc`` (group «شرکا»)
   so a partner is a first-class party everywhere: weighbridge, payments,
   bank transactions, ledgers, trial balance and the dashboard.
2. ``00_Master ← Op`` → two new weighbridge operations:
       «خرید شراکتی (ورود)»   ورود · خرید · weight · partner بستانکار / X004 بدهکار
       «فروش شراکتی (خروج)»   خروج · فروش · weight · partner بدهکار / X006 بستانکار
3. ``01_Entry ← Sync`` → the account block grows to row 279 (SH01…SH20) and a
   ``ShrkTbl`` block (code · name · partner% · our%) is mirrored for lookups.
4. ``01_Entry ← Scale`` → six visible columns **AQ…AV** (نسبت شرکت، سهم شرکت
   ریال، سهم شریک ریال، سهم شرکت kg، سهم شریک kg، کنترل) and every ledger leg
   (V/W/Y/Z) is multiplied by the company ratio — so only our share of a
   partnership load ever reaches the journal.
5. ``01_Entry ← Shrk`` → the partner ledger: per-partner weights, values,
   our share, partner share, payments made to the partner and the remaining
   balance, plus a 150-row detail of every partnership weighbridge line.
6. ``02_Bank`` → two new transaction types (پرداخت سهم شراکت / دریافت از شریک).

Everything is driven from the mother file: the partner codes, names and
ratios all come from ``00_Master ← Shrk``.
"""
import os
import re

import openpyxl
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

import grow
import look as L

FA = L.FA
FA_T = L.FA_T
EN = L.EN

# ------------------------------------------------------------------ partners
# code, name, partner share %, subject, product, phone, mobile, bank, note
PARTNERS = [
    ("SH01", "سعید بابایی", 50, "کیوی — شراکت فصلی", "کیوی",
     "011-4221-3301", "0911-112-3344", "BN02", "شریک اصلی کیوی؛ خرید و فروش مشترک"),
    ("SH02", "رحمت‌الله اکبری", 40, "پرتقال تامسون", "پرتقال",
     "011-4221-3302", "0911-223-4455", "BN01", "شراکت در خرید تامسون باغ‌های شرق"),
    ("SH03", "مهدی رستمی", 35, "نارنگی پیش‌رس", "نارنگی",
     "011-4221-3303", "0911-334-5566", "BN03", "نارنگی؛ تسویه ماهانه"),
    ("SH04", "حسین نیک‌پور", 50, "کیوی — سردخانه", "کیوی",
     "011-4221-3304", "0911-445-6677", "BN02", "کیوی سردخانه‌ای"),
    ("SH05", "کریم صادقی", 30, "پرتقال خونی", "پرتقال",
     "011-4221-3305", "0911-556-7788", "BN04", ""),
    ("SH06", "عباس موسوی", 45, "نارنگی — صادراتی", "نارنگی",
     "011-4221-3306", "0911-667-8899", "BN01", "بار صادراتی مشترک"),
    ("SH07", "رضا کاظمی", 50, "کیوی", "کیوی",
     "011-4221-3307", "0911-778-9900", "BN03", ""),
    ("SH08", "مجید قاسمی", 40, "پرتقال والنسیا", "پرتقال",
     "011-4221-3308", "0911-889-0011", "BN02", ""),
    ("SH09", "علی شریفی", 25, "نارنج", "نارنج",
     "011-4221-3309", "0911-990-1122", "BN04", "نارنج؛ شراکت کوچک"),
    ("SH10", "ناصر رحیمی", 50, "کیوی — درجه یک", "کیوی",
     "011-4221-3310", "0911-101-2233", "BN01", ""),
    ("SH11", "بابک سلطانی", 35, "پرتقال — میدان بار", "پرتقال",
     "011-4221-3311", "0911-212-3344", "BN03", ""),
    ("SH12", "جواد کریمی", 60, "کیوی — صادراتی", "کیوی",
     "011-4221-3312", "0911-323-4455", "BN02", "سهم شریک ۶۰٪ — قرارداد کتبی"),
]
BLANK = [f"SH{i:02d}" for i in range(len(PARTNERS) + 1, 21)]
PARTNERS_ALL = PARTNERS + [(c, "", "", "", "", "", "", "", "") for c in BLANK]

OP_IN = ("OP20", "خرید شراکتی (ورود)")
OP_OUT = ("OP21", "فروش شراکتی (خروج)")

TITLE_IN = OP_IN[1]
TITLE_OUT = OP_OUT[1]

ACC_FIRST = 261          # Acc!A261 … A280  (the sheet is used up to row 260)
ACC_LAST = ACC_FIRST + 19          # 280
SYNC_FIRST = 260         # 01_Entry ← Sync!B260 … B279
SYNC_LAST = SYNC_FIRST + 19
PARTNER_COL = 97         # Sync!CS
DETAIL_ROWS = 150






def _def(wb, name, ref):
    """openpyxl equivalent of xlsxwriter's define_name()."""
    if name in wb.defined_names:
        del wb.defined_names[name]
    wb.defined_names.add(DefinedName(name, attr_text=ref.lstrip("=")))


def _dvs(ws, ref, spec):
    """xlsxwriter-style spec → openpyxl DataValidation."""
    kw = {"allow_blank": True, "showDropDown": False}
    kind = spec.get("validate")
    if kind == "list":
        kw.update(type="list", formula1=spec.get("source", "").lstrip("="))
        kw["showErrorMessage"] = True
    elif kind == "decimal":
        kw.update(type="decimal", operator=spec.get("criteria", "between"),
                  formula1=spec.get("minimum", 0), formula2=spec.get("maximum", 100))
        kw["showErrorMessage"] = True
    for src, dst in (("errorTitle", "errorTitle"), ("error", "error"),
                     ("promptTitle", "promptTitle"), ("prompt", "prompt")):
        if src in spec:
            kw[dst] = spec[src]
            kw["showErrorMessage" if src.startswith("error") else "showInputMessage"] = True
    dv = DataValidation(**kw)
    ws.add_data_validation(dv)
    dv.add(ref)
    return dv


def _box(color=L.LINE):
    s = Side(style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)


# ===================================================================== master
def _master_prd(wb):
    """کیوی must exist: it is the product the partnership is about."""
    ws = wb["Prd"]
    have = {str(ws.cell(r, 2).value or "").strip() for r in range(10, 30)}
    add = [(c, n, cat, "کیلوگرم", "BK03", note)
           for c, n, cat, note in (("PB16", "کیوی", "کیوی",
                                    "بار شاخص فصل سرد؛ شراکت با سعید بابایی"),
                                   ("PB17", "پرتقال والنسیا", "پرتقال",
                                    "رقم دیررس؛ شراکت با مجید قاسمی")) if n not in have]
    made = 0
    for code, name, cat, unit, basket, note in add:
        row = next((r for r in range(10, 30) if not ws.cell(r, 2).value), None)
        if row is None:
            break
        tpl = {c: ws.cell(row - 1, c)._style for c in range(1, 8)}
        for c, v in enumerate([code, name, cat, unit, basket, note], start=1):
            cell = ws.cell(row, c, v)
            cell._style = tpl[c]
        ws.cell(row, 7).value = f'=IF($B{row}="","",$B{row}&" – "&$C{row})'
        made += 1
    return made


def _master_grp(wb):
    """«شرکا» becomes an account group (reports, trial balance, dashboard)."""
    ws = wb["Grp"]
    used = [r for r in range(10, 60) if ws.cell(r, 1).value]
    names = {str(ws.cell(r, 1).value).strip() for r in used}
    if "شرکا" in names:
        return False
    r = (max(used) + 1) if used else 10
    vals = ["شرکا", "شراکت", "SH", "Partners", "rose"]
    for i, v in enumerate(vals, start=1):
        c = ws.cell(r, i, v)
        c.font = Font(name=FA, size=9.5, color="FF2B2A26")
        c.alignment = Alignment(horizontal="center", vertical="center", readingOrder=2)
        c.fill = PatternFill("solid", fgColor="FFFBF8F1")
        c.border = _box()
    ws.print_area = f"A1:E{max(used) + 1}"
    return True


def _master_acc(wb):
    """Append SH01…SH20 to the chart of accounts."""
    ws = wb["Acc"]
    used = [r for r in range(10, 320) if ws.cell(r, 1).value]
    codes = {str(ws.cell(r, 1).value) for r in used}
    if "SH01" in codes:
        return 0
    rows = used[-1] + 1
    tpl = {}
    for c in range(1, 12):
        tpl[c] = (ws.cell(used[-1], c)._style)
    made = 0
    for i, (code, name, ratio, subject, product, phone, mobile, bank, note) in enumerate(PARTNERS_ALL):
        r = rows + i
        vals = [code, "شرکا", "شراکت", name, phone, mobile, "—", 0,
                "طرف حساب", "بله", None,
                f"{subject} · نسبت شریک {ratio}٪" if name else ""]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(r, c, v)
            if c in tpl:
                cell._style = tpl[c]
        ws.cell(r, 11).value = (f'=IF($A{r}="","",IF(LEFT($A{r},2)="SH","شرکا",""))')
        ws.cell(r, 8).value = 0
        made += 1
    ws.print_area = f"A1:M{rows + made + 1}"
    return made


def _master_op(wb):
    """The two partnership weighbridge operations."""
    ws = wb["Op"]
    have = {str(ws.cell(r, 2).value).strip() for r in range(10, 34) if ws.cell(r, 2).value}
    if TITLE_IN in have:
        return 0
    rows = [
        (OP_IN[0], TITLE_IN, "ورود", "خرید", "بله", "بله", "وزن", "بستانکار", "X004",
         "بدهکار", "افزودن", "—", None, None, None, "بله", "rose",
         "بار شراکتی که وارد انبار می‌شود. طرف حساب = شریک (کد SHxx). مبلغ = وزن خالص × فی؛ "
         "در دفتر فقط «سهم شرکت» ثبت می‌شود: بدهکار موجودی مرکبات / بستانکار حساب شریک. "
         "پرداخت‌های بعدی به شریک، همان سهم ما را تسویه می‌کند."),
        (OP_OUT[0], TITLE_OUT, "خروج", "فروش", "بله", "بله", "وزن", "بدهکار", "X006",
         "بستانکار", "هیچ", "—", None, None, None, "بله", "sky",
         "فروش سهم شرکت از بار شراکتی. طرف حساب = خریدار. در دفتر فقط سهم شرکت ثبت می‌شود: "
         "بدهکار خریدار / بستانکار فروش مرکبات. وزن کامل بار از انبار خارج می‌شود."),
    ]
    made = 0
    for i, vals in enumerate(rows):
        r = 29 + i
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(r, c, v)
            cell._style = ws.cell(10, c)._style
        made += 1
    ws.print_area = f"A1:S{max(33, 29 + made)}"
    # refresh the little counters on top of the sheet
    ws["A6"] = '=COUNTIF($B$10:$B$33,"?*")'
    ws["E6"] = '=COUNTIF($S$10:$S$33,"بله")'
    ws["I6"] = '=COUNTIF($P$10:$P$33,"بله")'
    return made


def _master_shrk(wb):
    """The partner register sheet itself."""
    if "Shrk" in wb.sheetnames:
        del wb["Shrk"]
    ws = wb.create_sheet("Shrk", 5)
    last = 11
    L.brand(ws, last, "شراکت — شرکا و نسبت سهم",
            "هر شریک یک کد SHxx، یک نام و یک «نسبت سهم شریک» دارد  ·  "
            "نسبت شرکت خودکار = ۱۰۰ منهای آن  ·  همین جدول، کشوی باسکول و پرداخت‌ها را می‌سازد",
            icon="◈")
    L.kpis(ws, 6, last, [
        ("شرکای تعریف‌شده", '=COUNTIF($B$12:$B$31,"?*")', L.INT, L.FOREST),
        ("شرکای فعال", '=COUNTIFS($B$12:$B$31,"?*",$J$12:$J$31,"فعال")', L.INT, L.GREEN),
        ("میانگین نسبت شریک", '=IFERROR(ROUND(AVERAGE($C$12:$C$31),0),0)', '0"٪"', L.GOLD),
        ("جمع نسبت شرکت", '=SUM($D$12:$D$31)', '0"٪"', L.TERRA),
        ("موضوع‌های شراکت", '=COUNTIF($E$12:$E$31,"?*")', L.INT, L.BLUE),
    ])
    L.section(ws, 9, last, "۰۱", "فهرست شرکا", "کدها را تغییر ندهید؛ نام و نسبت را آزادانه به‌روز کنید")

    cols = [
        ("کد شریک", "Code", 9,
         "کد ثابت شریک (SH01…SH20). همین کد در باسکول، شیت پرداخت و تراکنش بانکی انتخاب می‌شود."),
        ("نام شریک", "Partner", 24,
         "نام کامل شخص/شرکت شریک. اگر ردیفی بی‌نام بماند، در کشوها دیده نمی‌شود."),
        ("نسبت سهم شریک ٪", "Partner %", 14,
         "درصد سهم شریک از بار و مبلغ شراکتی (۱ تا ۹۹). همین عدد، مبنای محاسبه سهم وزنی و ریالی است."),
        ("نسبت سهم شرکت ٪", "Our %", 14,
         "خودکار: ۱۰۰ منهای نسبت شریک. این نسبت، سهمی است که در دفتر روزنامه سند می‌خورد."),
        ("موضوع شراکت", "Subject", 26, "بار/محصول شراکتی و توضیح کوتاه قرارداد."),
        ("محصول", "Product", 14, "محصول اصلی این شراکت (کیوی، پرتقال، نارنگی …)."),
        ("تلفن", "Phone", 14, ""),
        ("همراه", "Mobile", 15, ""),
        ("بانک تسویه", "Bank", 11, "بانک شرکت که سهم شریک از آن پرداخت می‌شود (BNxx)."),
        ("وضعیت", "Status", 10, "فعال / غیرفعال."),
        ("توضیح", "Note", 34, "هر یادداشت دیگر: شماره قرارداد، شرایط تسویه …"),
    ]
    L.thead(ws, 10, cols)
    for i, spec in enumerate(cols, start=1):
        cell = ws.cell(11, i, spec[1])
        L.st(cell, name=EN, size=7.5, color=L.SAGE_L, bg=L.FOREST)

    for i, (code, name, ratio, subject, product, phone, mobile, bank, note) in enumerate(PARTNERS_ALL):
        r = 12 + i
        vals = [code, name, ratio if name else None, None, subject, product, phone,
                mobile, bank, "فعال" if name else None, note]
        for c, v in enumerate(vals, start=1):
            if v is None:
                continue
            cell = ws.cell(r, c, v)
            L.st(cell, name=EN if c in (1, 3, 9) else FA, size=10,
                 bg=L.INPUT if c in (2, 3, 5, 9, 10) else L.CREAM,
                 color=L.INK, h="center" if c in (1, 3, 9, 10) else "right",
                 fmt='0"٪"' if c == 3 else None)
        d = ws.cell(r, 4, f'=IF($C{r}="","",100-$C{r})')
        L.st(d, name=EN, size=10, bg=L.MINT, color=L.FOREST, fmt='0"٪"')
        if not name:
            for c in range(1, last + 1):
                ws.cell(r, c).fill = PatternFill("solid", fgColor="FFF7F1E6")
        ws.row_dimensions[r].height = 20

    z = 11 + len(PARTNERS_ALL)
    _dvs(ws, f"C12:C{z}",
         {"validate": "decimal", "criteria": "between", "minimum": 1, "maximum": 99,
          "errorTitle": "نسبت شریک",
          "error": "عدد بین ۱ و ۹۹ وارد کنید (درصد سهم شریک).",
          "promptTitle": "نسبت سهم شریک",
          "prompt": "مثلاً 50 یعنی نصف بار و مبلغ سهم شریک است و نصف دیگر سهم شرکت."})
    _dvs(ws, f"J12:J{z}", {"validate": "list", "source": '"فعال,غیرفعال"'})
    _dvs(ws, f"I12:I{z}", {"validate": "list", "source": "Bnk!$A$10:$A$19"})
    _dvs(ws, f"F12:F{z}", {"validate": "list", "source": "Prd!$B$10:$B$29"})

    L.note(ws, 12 + len(PARTNERS_ALL) + 1, last,
           "قاعده: سهم شریک از «وزن» محاسبه می‌شود. در باسکول، عملیات «خرید شراکتی (ورود)» یا "
           "«فروش شراکتی (خروج)» را بزنید و کد شریک را در ستون طرف حساب انتخاب کنید؛ سیستم وزن و "
           "مبلغ را به دو سهم تقسیم می‌کند و فقط سهم شرکت را سند می‌زند.", bg=L.MINT, h=42)
    L.note(ws, 12 + len(PARTNERS_ALL) + 2, last,
           "پرداخت به شریک = سهم واریزی ما. هر پرداخت به حساب SHxx در شیت پرداخت (01_Entry) یا "
           "تراکنش بانکی (02_Bank) ثبت شود تا مانده بدهی ما به شریک، خودکار محاسبه شود.",
           bg=L.SAND, h=38)
    L.page(ws, last, 12 + len(PARTNERS_ALL) + 2, landscape=True, header_rows="10:11")
    dw = {"A": 9, "B": 24, "C": 17, "D": 17, "E": 26, "F": 14, "G": 14, "H": 15,
          "I": 12, "J": 11, "K": 34}
    for col, w in dw.items():
        ws.column_dimensions[col].width = w
    return ws


def _master_names(wb):
    """ShrkTbl / ListShrk + grow the account ranges to row 279."""
    last = ACC_LAST              # 280: 250 accounts + 20 partners
    for name, dn in list(wb.defined_names.items()):
        if name not in ("AcctTbl", "AcctCodes", "ListAllCodes", "ListAcctNames"):
            continue
        txt = dn.attr_text
        new = re.sub(r"\$([A-Z]{1,3})\$10:\$([A-Z]{1,3})\$259",
                     lambda m: f"${m.group(1)}$10:${m.group(2)}${last}", txt)
        new = new.replace("$D$259", f"$D${last}").replace("$C$259", f"$C${last}")
        if new != txt:
            dn.attr_text = new
    _def(wb, "ShrkTbl", "=Shrk!$A$12:$D$31")
    _def(wb, "ListShrk",
                   '=OFFSET(Shrk!$A$12,0,0,MAX(COUNTIF(Shrk!$B$12:$B$31,"?*"),1),1)')
    _def(wb, "ListShrkNames",
                   '=OFFSET(Shrk!$B$12,0,0,MAX(COUNTIF(Shrk!$B$12:$B$31,"?*"),1),1)')


def _master_home(wb):
    ws = wb["Home"]
    r = 50
    for row in range(44, 60):
        if ws.cell(row, 1).value and "Shrk" in str(ws.cell(row, 1).value):
            return
    ws.merge_cells(f"A{r}:J{r}")
    c = ws.cell(r, 1, "     شیت تازه:  Shrk — شرکا و نسبت شراکت  (۲۰ شریک، نسبت سهم شریک و سهم شرکت)")
    L.st(c, name=FA, size=10.5, bold=True, color=L.WHITE, bg=L.TERRA, border=L.TERRA, h="right")
    c.hyperlink = "#'Shrk'!A1"
    ws.row_dimensions[r].height = 24
    ws.print_area = f"A1:J{r}"


def _master_help(wb):
    ws = wb["Help"]
    r = ws.max_row + 2
    blocks = [
        ("۰۹", "شراکت — خرید و فروش مشترک"),
        ("p", "شیت Shrk فهرست شرکاست: کد SHxx، نام، «نسبت سهم شریک» و «نسبت سهم شرکت» (خودکار ۱۰۰ منهای آن). "
              "همین کدها به شیت Acc هم اضافه شده‌اند (گروه «شرکا»)، پس در باسکول، شیت پرداخت، تراکنش بانکی، "
              "دفاتر و ترازنامه مثل هر طرف حساب دیگری انتخاب می‌شوند."),
        ("b", "در باسکول دو عملیات تازه اضافه شده: «خرید شراکتی (ورود)» و «فروش شراکتی (خروج)». "
              "کافی است عملیات را انتخاب و کد شریک را در ستون «کد حساب / مقصد وجه» بزنید."),
        ("b", "محاسبه به وزن است: شیت Scale شش ستون تازه دارد — نسبت شرکت، سهم شرکت (ریال)، سهم شریک (ریال)، "
              "سهم شرکت (kg)، سهم شریک (kg) و کنترل. مبلغ هر ردیف = وزن خالص × فی."),
        ("b", "در دفتر روزنامه فقط «سهم شرکت» سند می‌خورد (خرید: بدهکار موجودی / بستانکار شریک — فروش: "
              "بدهکار خریدار / بستانکار فروش). سهم شریک در گزارش شیت Shrk فایل 01_Entry دیده می‌شود."),
        ("b", "پرداخت به شریک = سهم واریزی ما: در شیت پرداخت نوع «پرداخت»، کد SHxx و روش پرداخت را بزنید. "
              "این پرداخت‌ها در جدول شرکا از بدهی ما کم می‌شود و ستون «مانده» بدهی باقی‌مانده را نشان می‌دهد."),
        ("b", "برای تغییر نسبت یک شریک، فقط ستون «نسبت سهم شریک ٪» را در همین شیت عوض کنید؛ فایل‌های دیگر "
              "با باز شدن، Sync خود را از نو می‌خوانند (نسخه جدید پوشه، همان کدها)."),
    ]
    for kind, text in blocks:
        if kind == "۰۹":
            ws.merge_cells(f"A{r}:J{r}")
            c = ws.cell(r, 1, "۰۹   " + text)
            L.st(c, name=FA, size=11, bold=True, color=L.WHITE, bg=L.FOREST, border=L.FOREST, h="right")
            ws.row_dimensions[r].height = 26
        else:
            ws.merge_cells(f"A{r}:J{r}")
            c = ws.cell(r, 1, ("     •  " if kind == "b" else "") + text)
            L.st(c, name=FA, size=10, color=L.INK, bg=L.MINT if kind == "b" else L.CREAM,
                 border=L.LINE_S, h="right", wrap=True)
            ws.row_dimensions[r].height = 46
        r += 1
    ws.print_area = f"A1:J{r - 1}"


def master(path):
    """Everything that changes in 00_Master.xlsx."""
    wb = openpyxl.load_workbook(path)
    grp = _master_grp(wb)
    prd = _master_prd(wb)
    acc = _master_acc(wb)
    ops = _master_op(wb)
    _master_shrk(wb)
    _master_names(wb)
    _master_home(wb)
    _master_help(wb)
    wb.calculation.fullCalcOnLoad = True
    wb.active = 0
    wb.save(path)
    wb.close()
    return {"grp": grp, "prd": prd, "acc": acc, "ops": ops}


def read_partners(master_path):
    """(code, name, partner%, our%) straight from 00_Master ← Shrk."""
    wb = openpyxl.load_workbook(master_path, data_only=False)
    ws = wb["Shrk"]
    out = []
    for i in range(20):
        r = 12 + i
        code = str(ws.cell(r, 1).value or "").strip()
        name = str(ws.cell(r, 2).value or "").strip()
        ratio = ws.cell(r, 3).value
        try:
            ratio = float(ratio)
        except (TypeError, ValueError):
            ratio = 0.0
        out.append((code, name, ratio, round(100 - ratio, 4) if name else 0.0))
    wb.close()
    return out


# =================================================================== 01_Entry
def _sync(wb, partners):
    """Accounts block grows by the partners; ShrkTbl mirrored at CS…CV."""
    ws = wb["Sync"]
    for i, (code, name, ratio, our) in enumerate(partners):
        r = SYNC_FIRST + i
        vals = [code, name, "شرکا", "شراکت", "طرف حساب", 0]
        for c, v in enumerate(vals, start=2):
            cell = ws.cell(r, c, v)
            cell._style = ws.cell(SYNC_FIRST - 1, c)._style

    first, last = 8, 8 + len(partners) + 1          # header rows 8/9, data 10…29
    heads = [("شرکا و نسبت سهم", 4, L.TERRA),
             ("کد شریک", 1, L.TERRA), ("نام شریک", 1, L.TERRA),
             ("نسبت شریک ٪", 1, L.TERRA), ("نسبت شرکت ٪", 1, L.TERRA)]
    ws.merge_cells(start_row=first, start_column=PARTNER_COL,
                   end_row=first, end_column=PARTNER_COL + 3)
    c = ws.cell(first, PARTNER_COL, "شرکا و نسبت سهم  ·  PARTNERS")
    c.font = Font(name=FA, size=9.5, bold=True, color="FFFFFFFF")
    c.fill = PatternFill("solid", fgColor="FFC4623F")
    c.alignment = Alignment(horizontal="center", vertical="center", readingOrder=2)
    for j, h in enumerate(["کد", "نام شریک", "نسبت شریک ٪", "نسبت شرکت ٪"]):
        cell = ws.cell(first + 1, PARTNER_COL + j, h)
        cell.font = Font(name=FA, size=9, bold=True, color="FFFFFFFF")
        cell.fill = PatternFill("solid", fgColor="FFC4623F")
        cell.alignment = Alignment(horizontal="center", vertical="center", readingOrder=2)
        cell.border = _box("FFC4623F")
    for i, (code, name, ratio, our) in enumerate(partners):
        r = first + 2 + i
        for j, v in enumerate([code, name, ratio if name else 0, our]):
            cell = ws.cell(r, PARTNER_COL + j, v)
            cell.font = Font(name=EN if j in (0, 2, 3) else FA, size=9,
                             color="FF2B2A26")
            cell.alignment = Alignment(horizontal="right" if j == 1 else "center",
                                       vertical="center", readingOrder=2)
            cell.border = _box("FFDDD5C4")
            if j in (2, 3):
                cell.number_format = '0.0"٪"'

    _def(wb, "ShrkTbl", f"=Sync!$CS$10:$CV${first + 1 + len(partners)}")
    _def(wb, "ListShrk",
                   '=OFFSET(Sync!$CS$10,0,0,MAX(COUNTIF(Sync!$CT$10:$CT$29,"?*"),1),1)')
    # account ranges → row 279
    for name, dn in list(wb.defined_names.items()):
        if name not in ("AcctTbl", "AcctCodes", "ListAllCodes", "ListAcctNames"):
            continue
        txt = dn.attr_text
        new = re.sub(r"\$([A-Z]{1,3})\$10:\$([A-Z]{1,3})\$259",
                     lambda m: f"${m.group(1)}$10:${m.group(2)}${SYNC_LAST}", txt)
        new = new.replace("$C$259", f"$C${SYNC_LAST}")
        if new != txt:
            dn.attr_text = new
    return ws


def _sync_ops(wb, master_path):
    """Copy OP20 / OP21 (master Op columns B…P) into the Sync operation table."""
    mwb = openpyxl.load_workbook(master_path, read_only=True)
    mop = mwb["Op"]
    rows = []
    for r in range(10, 34):
        title = str(mop.cell(r, 2).value or "").strip()
        if title in (TITLE_IN, TITLE_OUT):
            rows.append([mop.cell(r, c).value for c in range(2, 17)])
    mwb.close()
    ws = wb["Sync"]
    tgt = 29
    for i, vals in enumerate(rows):
        r = tgt + i
        for c, v in enumerate(vals, start=9):
            cell = ws.cell(r, c, v)
            try:
                cell._style = ws.cell(28, c)._style
            except Exception:
                pass
    return len(rows)


def scale_columns(ws):
    """AQ…AW (+ two hidden helpers): the weight-based partnership share.

    AQ کد شریک      the partner of this load (input; optional on purchases —
                     it falls back to the party code in column G)
    AR نسبت شرکت    our ratio, 100 − partner share, read from ShrkTbl
    AS سهم شرکت     our share of the money  (what reaches the journal)
    AT سهم شریک     the partner's share of the money
    AU سهم شرکت kg  our share of the net weight  ← the weight-based split
    AV سهم شریک kg  the partner's share of the weight
    AW کنترل        ✓ / ⚠
    AX (hidden)     the resolved partner code, so the report can SUMIFS on it
    AY (hidden)     running number of partnership rows (fast detail lookup)
    """
    last = grow.SCALE_LAST
    cols = ["AQ", "AR", "AS", "AT", "AU", "AV", "AW"]
    heads = [
        ("کد شریک", "Partner",
         "کد شریک (SHxx). در «خرید شراکتی» می‌توانید خالی بگذارید تا همان کد طرف حساب به کار برود؛ "
         "در «فروش شراکتی» حتماً شریک را اینجا انتخاب کنید (ستون طرف حساب، خریدار است)."),
        ("نسبت شرکت ٪", "Our %",
         "سهم شرکت از این بار شراکتی (۱۰۰ منهای نسبت شریک، از 00_Master ← Shrk). "
         "برای عملیات غیرشراکتی ۱۰۰٪ است."),
        ("سهم شرکت — ریال", "Our rial",
         "مبلغ کل ردیف × نسبت شرکت. همین عدد وارد دفتر روزنامه می‌شود."),
        ("سهم شریک — ریال", "Partner rial",
         "مبلغ کل ردیف × نسبت شریک؛ فقط در گزارش شراکت دیده می‌شود."),
        ("سهم شرکت — کیلو", "Our kg",
         "وزن خالص × نسبت شرکت — مبنای محاسبه سهم در این نوع عملیات."),
        ("سهم شریک — کیلو", "Partner kg",
         "وزن خالص × نسبت شریک."),
        ("کنترل شراکت", "Check",
         "✓ سهمی = نسبت شریک پیدا شد و محاسبه انجام شد.  ⚠ نسبت شریک = کد شریک در جدول شرکا نیست "
         "یا نسبتش صفر است؛ ابتدا نسبت را در 00_Master ← Shrk وارد کنید."),
    ]
    ws.merge_cells("AQ7:AW7")
    band = ws["AQ7"]
    band.value = "شراکت  ·  سهم وزنی و ریالی  (PARTNERSHIP SHARE)"
    band.font = Font(name=FA, size=10, bold=True, color="FFFFFFFF")
    band.fill = PatternFill("solid", fgColor="FFC4623F")
    band.alignment = Alignment(horizontal="center", vertical="center", readingOrder=2)
    for c in range(43, 50):                        # AQ…AW
        ws.cell(7, c).fill = PatternFill("solid", fgColor="FFC4623F")

    for i, (col, fa, en, tip) in enumerate(zip(cols, [h[0] for h in heads],
                                               [h[1] for h in heads],
                                               [h[2] for h in heads])):
        cell = ws[f"{col}8"]
        try:
            cell._style = ws["T8"]._style
        except Exception:
            pass
        cell.value = fa
        L.tip(ws, f"{col}8", tip, w=340, h=150)
        ws.row_dimensions[8].height = max(ws.row_dimensions[8].height or 0, 42)
        ws.row_dimensions[7].height = max(ws.row_dimensions[7].height or 0, 20)
        e = ws[f"{col}9"]
        try:
            e._style = ws["T9"]._style
        except Exception:
            pass
        e.value = en
        ws.column_dimensions[col].width = 15
        for r in (8, 9):
            ws[f"{col}{r}"].fill = PatternFill("solid", fgColor="FFC4623F")
            ws[f"{col}{r}"].font = Font(name=FA if r == 8 else EN, size=9, bold=True,
                                        color="FFFFFFFF")
            ws[f"{col}{r}"].alignment = Alignment(horizontal="center", vertical="center",
                                                  wrap_text=True, readingOrder=2)

    f = {
        "AR": ('=IF($E{r}="",1,IF(ISNUMBER(SEARCH("شراکتی",$E{r})),'
               'IFERROR(VLOOKUP($AX{r},ShrkTbl,4,0),0)/100,1))'),
        "AS": ('=IF(OR($AJ{r}="",$AR{r}<=0,$AR{r}>=1),"",ROUND($AJ{r}*$AR{r},0))'),
        "AT": ('=IF(OR($AJ{r}="",$AR{r}<=0,$AR{r}>=1),"",'
               'ROUND($AJ{r}*(1-$AR{r}),0))'),
        "AU": ('=IF(OR($N{r}="",$AR{r}<=0,$AR{r}>=1),"",'
               'ROUND($N{r}*$AR{r},2))'),
        "AV": ('=IF(OR($N{r}="",$AR{r}<=0,$AR{r}>=1),"",'
               'ROUND($N{r}*(1-$AR{r}),2))'),
        "AW": ('=IF($E{r}="","",IF(ISNUMBER(SEARCH("شراکتی",$E{r})),'
               'IF($AR{r}<=0,"⚠ نسبت شریک",IF($AQ{r}="","✓ سهمی (طرف حساب)","✓ سهمی")),"—"))'),
        "AX": '=IF($AQ{r}="",$U{r},$AQ{r})',
        "AY": ('=IF(OR($E{r}="",NOT(ISNUMBER(SEARCH("شراکتی",$E{r})))),"",'
               'COUNT($AY$9:AY{r-1})+1)'),
    }
    for r in range(10, last + 1):
        # the input column: styled everywhere, value only where the user typed one
        cell = ws[f"AQ{r}"]
        cell.border = _box("FFEDE6D8")
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=False, readingOrder=2)
        cell.fill = PatternFill("solid", fgColor="FFFFF4D6")
        cell.font = Font(name=EN, size=8.5, bold=True, color="FF2B2A26")
        for col, tpl in f.items():
            cell = ws[f"{col}{r}"]
            if col == "AY":
                cell.value = tpl.replace("{r-1}", str(r - 1)).format(r=r)
            else:
                cell.value = tpl.format(r=r)
            cell.border = _box("FFEDE6D8")
            cell.alignment = Alignment(horizontal="center", vertical="center",
                                       wrap_text=False, readingOrder=2)
            cell.font = Font(name=EN if col not in ("AW",) else FA, size=8.5,
                             color="FF2B2A26")
            if col == "AQ":
                cell.fill = PatternFill("solid", fgColor="FFFFF4D6")
                cell.font = Font(name=EN, size=8.5, bold=True, color="FF2B2A26")
            elif col == "AR":
                cell.number_format = "0.0%"
                cell.fill = PatternFill("solid", fgColor="FFE4F0E8")
            elif col in ("AS", "AT"):
                cell.number_format = "#,##0"
                cell.fill = PatternFill("solid", fgColor="FFFBF8F1")
            elif col in ("AU", "AV"):
                cell.number_format = "#,##0.00"
                cell.fill = PatternFill("solid", fgColor="FFFBF8F1")
            elif col == "AW":
                cell.fill = PatternFill("solid", fgColor="FFE4F0E8")
                cell.font = Font(name=FA, size=8.5, bold=True, color="FF3F7A57")
            else:                                   # AX / AY helpers
                cell.font = Font(name=EN, size=8, color="FF6B7A76")

    ws.column_dimensions["AX"].hidden = True
    ws.column_dimensions["AY"].hidden = True
    _dvs(ws, f"AQ10:AQ{last}", {"validate": "list", "source": "ListShrk"})

    # partnership rows keep a tinted body so they stand out in the printed sheet
    ws.conditional_formatting.add(
        f"A10:AW{last}",
        FormulaRule(formula=['ISNUMBER(SEARCH("شراکتی",$E10))'],
                    fill=PatternFill("solid", bgColor="FFFDF3E4"), stopIfTrue=False))
    ws.conditional_formatting.add(
        f"AW10:AW{last}",
        FormulaRule(formula=['$AW10="⚠ نسبت شریک"'],
                    fill=PatternFill("solid", bgColor="FFF2DCD3"),
                    font=Font(name=FA, size=8.5, bold=True, color="FFB4544A")))
    return len(cols)


def scale_legs(ws):
    """Multiply the two journal legs by the company ratio (partnership only)."""
    last = grow.SCALE_LAST
    n = 0
    for col in ("V", "W", "Y", "Z"):
        for r in range(10, last + 1):
            cell = ws[f"{col}{r}"]
            v = cell.value
            if isinstance(v, str) and v.endswith(f',"",$AJ{r})'):
                cell.value = v[:-1] + f"*$AR{r})"
                n += 1
    return n


def scale_hide_helpers(ws):
    """The twelve engineering columns (U…AP) move behind the curtain."""
    for c in range(21, 43):
        ws.column_dimensions[get_column_letter(c)].hidden = True
    return 22


SCALE_TIPS = {
    "AQ": "کد شریک (SHxx). در خرید شراکتی اختیاری است (پیش‌فرض: همان کد طرف حساب)؛ "
          "در فروش شراکتی الزامی است چون ستون طرف حساب، خریدار است.",
    "AR": "نسبت سهم شرکت: ۱۰۰ − نسبت شریک، از 00_Master ← Shrk.",
    "AS": "سهم شرکت از مبلغ این ردیف (ریال) — همان مبلغی که در دفتر روزنامه ثبت می‌شود.",
    "AT": "سهم شریک از مبلغ این ردیف (ریال) — فقط در گزارش شراکت می‌آید.",
    "AU": "سهم شرکت از وزن خالص (کیلوگرم) — مبنای محاسبه سهم.",
    "AV": "سهم شریک از وزن خالص (کیلوگرم).",
    "AW": "✓ یعنی نسبت شریک پیدا شد؛ ⚠ یعنی کد شریک/نسبت در جدول شرکا ناقص است.",
}


# ------------------------------------------------------------------ demo data
# date, time, doc, op, party, product, basket, scale kg, baskets, price, note
DEMO_SCALE = [
    ("1405/07/14", "08:10", "BS-101", TITLE_IN, "SH01", "کیوی", "BK03 – ۹ک مشکی",
     5200, 260, 45000, "", "خرید مشترک کیوی با سعید بابایی"),
    ("1405/07/16", "09:25", "BS-102", TITLE_IN, "SH06", "نارنگی پیچ", "BK05 – تک‌رج آبی بزرگ",
     3100, 150, 32000, "", "خرید مشترک نارنگی با عباس موسوی"),
    ("1405/07/18", "11:40", "BS-103", TITLE_IN, "SH02", "پرتقال تامسون", "BK01 – ۶ک مشکی",
     2400, 120, 38000, "", "خرید مشترک تامسون با رحمت‌الله اکبری"),
    ("1405/07/19", "16:40", "BS-104", TITLE_OUT, "SL002", "کیوی", "BK03 – ۹ک مشکی",
     5000, 250, 62000, "SH01", "فروش سهم شرکت از کیوی شراکتی"),
    ("1405/07/21", "07:55", "BS-105", TITLE_OUT, "SL001", "نارنگی پیچ", "BK05 – تک‌رج آبی بزرگ",
     3000, 145, 41000, "SH06", "فروش سهم شرکت از نارنگی شراکتی"),
    ("1405/07/22", "13:20", "BS-106", TITLE_OUT, "SL002", "پرتقال تامسون", "BK01 – ۶ک مشکی",
     2300, 115, 44000, "SH02", "فروش سهم شرکت از تامسون شراکتی"),
]

# date, doc, type, method, party, bank, cheque, amount, cheque status, due, note
DEMO_PAY = [
    ("1405/07/18", "PV-101", "پرداخت", "انتقال بانکی", "SH01", "BN02", "", 60000000,
     "", "", "پرداخت سهم ما به شریک — کیوی"),
    ("1405/07/20", "PV-102", "پرداخت", "چک", "SH06", "BN01", "CHQ-1050", 40000000,
     "در جریان", "1405/08/20", "سهم ما از نارنگی شراکتی"),
    ("1405/07/23", "PV-103", "پرداخت", "نقد", "SH02", "", "", 15000000,
     "", "", "سهم ما از تامسون شراکتی"),
]


def demo_rows(wb):
    """Sample partnership loads + the payments that settle our share."""
    made = 0
    sc = wb["Scale"]
    last = grow.SCALE_LAST
    r0 = next((r for r in range(10, min(last, 400) + 1) if not sc.cell(r, 2).value), None)
    if r0 and r0 + len(DEMO_SCALE) - 1 <= last:
        for i, (day, tm, doc, op, party, prod, basket, kg, nb, price, partner,
                note) in enumerate(DEMO_SCALE):
            r = r0 + i
            for col, v in (("B", day), ("C", tm), ("D", doc), ("E", op), ("G", party),
                           ("J", prod), ("K", basket), ("L", kg), ("M", nb), ("O", price),
                           ("Q", 0), ("AQ", partner), ("T", note)):
                cell = sc[f"{col}{r}"]
                cell.value = v
                cell._style = sc[f"{col}{r0 - 1}"]._style
            made += 1
    pay = wb["Pay"]
    lastp = grow.PAY_LAST
    p0 = next((r for r in range(10, min(lastp, 200) + 1) if not pay.cell(r, 2).value), None)
    if p0 and p0 + len(DEMO_PAY) - 1 <= lastp:
        for i, (day, doc, kind, method, party, bank, chq, amount, st, due, note) in enumerate(DEMO_PAY):
            r = p0 + i
            for col, v in (("B", day), ("C", doc), ("D", kind), ("E", method), ("F", party),
                           ("I", bank), ("J", chq), ("K", amount), ("L", st), ("M", due),
                           ("P", note)):
                cell = pay[f"{col}{r}"]
                cell.value = v
                cell._style = pay[f"{col}{p0 - 1}"]._style
            made += 1
    return made


# ============================================================== partner ledger
def report(wb, partners):
    """Sheet ``Shrk`` in 01_Entry — the partner ledger."""
    if "Shrk" in wb.sheetnames:
        del wb["Shrk"]
    ws = wb.create_sheet("Shrk", wb.sheetnames.index("Pay") + 1)
    n_p = 20
    last = 14
    S = grow.SCALE_LAST
    B = grow.BOOK_LAST
    head1 = 14                       # first partner row
    rows = head1 + n_p               # total row
    s2 = rows + 2                    # section ۲
    th2 = s2 + 1
    d1 = s2 + 2

    L.brand(ws, last, "حساب شرکا — سهم وزنی، سهم ریالی و مانده",
            "خرید و فروش شراکتی از باسکول خوانده می‌شود  ·  سهم شرکت از «وزن» محاسبه می‌شود  ·  "
            "پرداخت به شریک = سهم واریزی ما", icon="◈")

    L.kpis(ws, 6, last, [
        ("وزن خرید شراکتی (kg)", f"=SUM($E${head1}:$E${rows - 1})", "#,##0", L.FOREST),
        ("سهم شرکت از خرید", f"=SUM($G${head1}:$G${rows - 1})", L.INT, L.TERRA),
        ("پرداختی ما به شرکا", f"=SUM($L${head1}:$L${rows - 1})", L.INT, L.BLUE),
        ("مانده بدهی ما", f"=SUM($M${head1}:$M${rows - 1})", L.INT, L.RED),
        ("سهم شرکت از فروش", f"=SUM($K${head1}:$K${rows - 1})", L.INT, L.GREEN),
    ])

    # ---- ۰۱ settings: the two operation titles
    ws.merge_cells("A9:B9")
    L.st(ws["A9"], name=FA, size=10, bold=True, color=L.FOREST, bg=L.SAND, border=L.LINE, h="right")
    ws["A9"] = "عملیات ورود شراکتی"
    ws.merge_cells("C9:E9")
    c = ws["C9"]
    c.value = TITLE_IN
    L.st(c, name=FA, size=10, bold=True, color=L.INK, bg=L.INPUT, border=L.LINE, h="center")
    ws.merge_cells("A10:B10")
    L.st(ws["A10"], name=FA, size=10, bold=True, color=L.FOREST, bg=L.SAND, border=L.LINE, h="right")
    ws["A10"] = "عملیات خروج شراکتی"
    ws.merge_cells("C10:E10")
    c = ws["C10"]
    c.value = TITLE_OUT
    L.st(c, name=FA, size=10, bold=True, color=L.INK, bg=L.INPUT, border=L.LINE, h="center")
    for cell in ("C9", "C10"):
        _dvs(ws, cell, {"validate": "list", "source": "ListOps"})
    ws.merge_cells("F9:N10")
    n = ws["F9"]
    n.value = ("همه فرمول‌های این شیت از همین دو خانه می‌خوانند؛ اگر عنوان عملیات را در فایل مادر "
               "عوض کردید، اینجا هم همان عنوان را انتخاب کنید.")
    L.st(n, name=FA, size=9, color=L.MUTED, bg=L.CREAM, border=L.LINE_S, h="right", wrap=True)
    ws.row_dimensions[9].height = 22
    ws.row_dimensions[10].height = 22

    cols = [
        ("کد شریک", "Code", 9, "کد ثابت شریک؛ همان کدی که در باسکول و پرداخت انتخاب می‌شود."),
        ("نام شریک", "Partner", 20, "نام از 00_Master ← Shrk."),
        ("نسبت شریک", "Partner %", 10, "درصد سهم شریک از بار و مبلغ."),
        ("نسبت شرکت", "Our %", 10, "۱۰۰ − نسبت شریک؛ همین نسبت در دفتر سند می‌خورد."),
        ("وزن خرید", "In kg", 12, "جمع وزن خالص ردیف‌های «خرید شراکتی» این شریک."),
        ("مبلغ خرید", "In value", 14, "جمع مبلغ کل خرید شراکتی (وزن × فی)."),
        ("سهم شرکت از خرید", "Our purchase", 14,
         "بستانکار حساب شریک — بدهی ما بابت سهم خودمان از خرید."),
        ("سهم شریک از خرید", "Partner purchase", 13, "سهم شریک از خرید؛ در دفتر شرکت نمی‌آید."),
        ("وزن فروش", "Out kg", 12, "جمع وزن خالص ردیف‌های «فروش شراکتی» این شریک."),
        ("مبلغ فروش", "Out value", 14, "جمع مبلغ کل فروش شراکتی."),
        ("سهم شرکت از فروش", "Our sales", 14, "بستانکار فروش — درآمد سهم خودمان."),
        ("پرداختی ما", "Paid", 13, "جمع بدهکار حساب شریک: پرداخت‌های ما به او."),
        ("مانده (بدهی ما)", "Balance", 14,
         "سهم شرکت از خرید − پرداختی‌ها. مثبت = ما به شریک بدهکاریم."),
        ("وضعیت", "Status", 11, "بدهی ما / تسویه / طلب ما / ⚠ نسبت."),
    ]
    L.section(ws, 12, last, "۰۱", "حساب هر شریک",
              "وزن، سهم و مانده — همه از باسکول و دفتر روزنامه")
    hdr = 13
    L.thead(ws, hdr, cols, height=40)

    for i, (code, name, ratio, our) in enumerate(partners):
        r = head1 + i
        ws.cell(r, 1, code)
        ws.cell(r, 2, f'=IF($A{r}="","",IFERROR(VLOOKUP($A{r},ShrkTbl,2,0),"کد ناشناخته"))')
        ws.cell(r, 3, f'=IF($A{r}="","",IFERROR(VLOOKUP($A{r},ShrkTbl,3,0),"")/100)')
        ws.cell(r, 4, f'=IF($A{r}="","",IFERROR(VLOOKUP($A{r},ShrkTbl,4,0),"")/100)')
        ws.cell(r, 5, f'=IF($A{r}="","",SUMIFS(Scale!$N$10:$N${S},Scale!$E$10:$E${S},$C$9,'
                      f'Scale!$AX$10:$AX${S},$A{r}))')
        ws.cell(r, 6, f'=IF($A{r}="","",SUMIFS(Scale!$AJ$10:$AJ${S},Scale!$E$10:$E${S},$C$9,'
                      f'Scale!$AX$10:$AX${S},$A{r}))')
        ws.cell(r, 7, f'=IF($A{r}="","",SUMIFS(Scale!$AS$10:$AS${S},Scale!$E$10:$E${S},$C$9,'
                      f'Scale!$AX$10:$AX${S},$A{r}))')
        ws.cell(r, 8, f'=IF($A{r}="","",$F{r}-$G{r})')
        ws.cell(r, 9, f'=IF($A{r}="","",SUMIFS(Scale!$N$10:$N${S},Scale!$E$10:$E${S},$C$10,'
                      f'Scale!$AX$10:$AX${S},$A{r}))')
        ws.cell(r, 10, f'=IF($A{r}="","",SUMIFS(Scale!$AJ$10:$AJ${S},Scale!$E$10:$E${S},$C$10,'
                       f'Scale!$AX$10:$AX${S},$A{r}))')
        ws.cell(r, 11, f'=IF($A{r}="","",SUMIFS(Scale!$AS$10:$AS${S},Scale!$E$10:$E${S},$C$10,'
                       f'Scale!$AX$10:$AX${S},$A{r}))')
        ws.cell(r, 12, f'=IF($A{r}="","",SUMIFS(Book!$L$10:$L${B},Book!$D$10:$D${B},$A{r}))')
        ws.cell(r, 13, f'=IF($A{r}="","",$G{r}-$L{r})')
        ws.cell(r, 14, (f'=IF($A{r}="","",IF(AND($D{r}=0,$E{r}+$I{r}>0),"⚠ نسبت",'
                        f'IF(ROUND($M{r},0)>0,"بدهی ما",IF(ROUND($M{r},0)<0,"طلب ما","تسویه"))))'))
        for c in range(1, last + 1):
            cell = ws.cell(r, c)
            cell.border = _box("FFEDE6D8")
            cell.alignment = Alignment(horizontal="center" if c not in (2, 14) else "center",
                                       vertical="center", readingOrder=2)
            cell.font = Font(name=EN if c in (1, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13) else FA,
                             size=9.5, color="FF2B2A26")
            if c in (3, 4):
                cell.number_format = "0%"
            elif c in (5, 9):
                cell.number_format = "#,##0"
            elif c in (6, 7, 8, 10, 11, 12, 13):
                cell.number_format = "#,##0"
            if c in (12, 13):
                cell.fill = PatternFill("solid", fgColor="FFFBF8F1")
            if not name:
                cell.fill = PatternFill("solid", fgColor="FFF7F1E6")
        ws.row_dimensions[r].height = 20

    # total row
    ws.merge_cells(f"A{rows}:D{rows}")
    t = ws.cell(rows, 1, "جمع")
    L.st(t, name=FA, size=10, bold=True, color=L.WHITE, bg=L.FOREST, h="center")
    for c in range(1, 5):
        ws.cell(rows, c).fill = PatternFill("solid", fgColor="FF14453D")
        ws.cell(rows, c).border = _box("FF14453D")
    for c in (5, 6, 7, 8, 9, 10, 11, 12, 13):
        cell = ws.cell(rows, c, f"=SUM({get_column_letter(c)}{head1}:{get_column_letter(c)}{rows - 1})")
        cell.number_format = "#,##0"
        cell.font = Font(name=EN, size=9.5, bold=True, color="FF14453D")
        cell.fill = PatternFill("solid", fgColor="FFE4F0E8")
        cell.border = _box("FF7FA99B")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    st = ws.cell(rows, 14, f'=IF(ROUND($M{rows},0)>0,"بدهی ما",IF(ROUND($M{rows},0)<0,"طلب ما","تسویه"))')
    L.st(st, name=FA, size=9.5, bold=True, color=L.FOREST, bg=L.MINT, border=L.SAGE)

    ws.conditional_formatting.add(
        f"N{head1}:N{rows}",
        FormulaRule(formula=[f'$N{head1}="بدهی ما"'],
                    font=Font(name=FA, size=9.5, bold=True, color="FFB4544A"),
                    fill=PatternFill("solid", bgColor="FFF2DCD3")))
    ws.conditional_formatting.add(
        f"N{head1}:N{rows}",
        FormulaRule(formula=[f'$N{head1}="تسویه"'],
                    font=Font(name=FA, size=9.5, bold=True, color="FF3F7A57"),
                    fill=PatternFill("solid", bgColor="FFE4F0E8")))

    L.note(ws, rows + 1, last,
           "مانده = سهم شرکت از خرید شراکتی − پرداخت‌های ما به شریک. «سهم شرکت از خرید» بستانکار "
           "حساب شریک است (بدهی ما) و هر پرداخت به کد SHxx آن را کم می‌کند. مانده مثبت یعنی هنوز "
           "به شریک بدهکاریم.", bg=L.CREAM, h=34)

    # ---- ۰۲ detail
    L.section(ws, s2, last, "۰۲", "ریز عملیات شراکتی",
              "تا ۱۵۰ ردیف آخر؛ کل ردیف‌ها در شیت Book است")
    det = [
        ("تاریخ", 10), ("شریک", 20), ("جهت", 9), ("محصول", 12), ("وزن (kg)", 11),
        ("فی", 12), ("مبلغ کل", 14), ("نسبت شرکت", 10), ("سهم شرکت", 14),
        ("سهم شریک", 13), ("وزن سهم شرکت", 12), ("سند", 10),
    ]
    for i, (fa, w) in enumerate(det, start=1):
        c = ws.cell(th2, i, fa)
        L.st(c, name=FA, size=8.5, bold=True, color=L.WHITE, bg=L.FOREST_2, border=L.FOREST_2,
             h="center", wrap=True)
        ws.column_dimensions[get_column_letter(i)].width = max(
            ws.column_dimensions[get_column_letter(i)].width or 0, w)
    ws.cell(th2, 15, "pos").font = Font(color="FFFFFFFF", size=7)
    ws.row_dimensions[th2].height = 26

    for k in range(DETAIL_ROWS):
        r = d1 + k
        ws.cell(r, 15, f'=IFERROR(MATCH({k + 1},Scale!$AY$10:$AY${S},0),"")')
        src = {
            1: "B", 2: "PARTNER", 3: "F", 4: "J", 5: "N", 6: "O", 7: "AJ", 8: "AR",
            9: "AS", 10: "AT", 11: "AU", 12: "D",
        }
        for c, col in src.items():
            if col == "PARTNER":        # partner name (falls back to the party name)
                fml = (f'=IF($O{r}="","",IFERROR(VLOOKUP(INDEX(Scale!$AX$10:$AX${S},$O{r}),'
                       f'ShrkTbl,2,0),INDEX(Scale!$H$10:$H${S},$O{r})))')
            else:
                fml = f'=IF($O{r}="","",INDEX(Scale!${col}$10:${col}${S},$O{r}))'
            cell = ws.cell(r, c, fml)
            cell.border = _box("FFEDE6D8")
            cell.alignment = Alignment(horizontal="center", vertical="center", readingOrder=2)
            cell.font = Font(name=EN if c not in (2, 3, 4) else FA, size=8.5, color="FF2B2A26")
            if c in (5, 7, 9, 10, 11):
                cell.number_format = "#,##0"
            if c == 8:
                cell.number_format = "0%"
        if k % 2:
            for c in range(1, 13):
                ws.cell(r, c).fill = PatternFill("solid", fgColor="FFFBF8F1")
    d_last = d1 + DETAIL_ROWS - 1
    L.note(ws, d_last + 2, last,
           "ردیف‌های این جدول خودکار از شیت Scale خوانده می‌شوند: هر ردیفی که عملیاتش یکی از دو "
           "عملیات شراکتی باشد. برای دیدن همه ردیف‌ها، در شیت Book ستون «عملیات / نوع» را روی "
           "«شراکتی» فیلتر کنید.", bg=L.MINT, h=36)

    ws.column_dimensions["O"].hidden = True
    L.page(ws, last, d_last + 2, landscape=True, header_rows=f"{hdr}:{hdr}")
    ws.auto_filter.ref = f"A{th2}:L{d_last}"
    return ws


def entry_help(wb):
    ws = wb["Help"]
    r = ws.max_row + 2
    blocks = [
        ("۱۱", "شراکت — خرید و فروش مشترک با شریک"),
        ("p", "وقتی باری (مثلاً کیوی) مشترک با یک شریک خریداری می‌شود، از دو عملیات تازه استفاده کنید: "
              "«خرید شراکتی (ورود)» برای ورود بار و «فروش شراکتی (خروج)» برای خروج آن. در هر دو، "
              "کد شریک (SHxx) یا خریدار را در ستون «کد حساب / مقصد وجه» انتخاب کنید و وزن و فی را بزنید."),
        ("b", "ستون‌های AQ تا AV شیت Scale سهم را نشان می‌دهند: نسبت شرکت، سهم شرکت (ریال)، سهم شریک (ریال)، "
              "سهم شرکت (کیلو) و سهم شریک (کیلو) — همه از «وزن» محاسبه می‌شوند. ستون کنترل می‌گوید نسبت "
              "شریک درست خوانده شده یا نه."),
        ("b", "در دفتر روزنامه فقط سهم شرکت می‌آید: خرید = بدهکار موجودی مرکبات / بستانکار حساب شریک؛ "
              "فروش = بدهکار خریدار / بستانکار فروش مرکبات. یعنی بار کامل در انبار ثبت می‌شود، ولی ارزش "
              "دفتری به اندازه سهم ماست."),
        ("b", "شیت Shrk همین فایل، حساب هر شریک را نشان می‌دهد: وزن و مبلغ خرید و فروش، سهم شرکت، سهم شریک، "
              "پرداخت‌های ما و مانده بدهی. پرداخت به شریک را در شیت Pay با نوع «پرداخت» و کد SHxx ثبت کنید."),
        ("b", "نسبت‌ها در 00_Master ← Shrk تعریف می‌شوند (کد، نام، نسبت سهم شریک). نسبت شرکت خودکار "
              "۱۰۰ منهای آن است. اگر کد شریک در آن جدول نباشد، ستون کنترل ⚠ می‌دهد و سهمی سند نمی‌خورد."),
    ]
    for kind, text in blocks:
        if kind == "۱۱":
            ws.merge_cells(f"A{r}:J{r}")
            c = ws.cell(r, 1, "۱۱   " + text)
            L.st(c, name=FA, size=11, bold=True, color=L.WHITE, bg=L.TERRA, border=L.TERRA,
                 h="right")
            ws.row_dimensions[r].height = 26
        else:
            ws.merge_cells(f"A{r}:J{r}")
            c = ws.cell(r, 1, "     •  " + text)
            L.st(c, name=FA, size=10, color=L.INK, bg=L.MINT, border=L.LINE_S, h="right",
                 wrap=True)
            ws.row_dimensions[r].height = 48
        r += 1
    ws.print_area = f"A1:J{r - 1}"       # last written row, not the empty one after it
    return r


def entry_home(wb):
    ws = wb["Home"]
    r = ws.max_row + 2
    ws.merge_cells(f"A{r}:J{r}")
    c = ws.cell(r, 1, "   شیت تازه:  Shrk — حساب شرکا (سهم وزنی، سهم ریالی، پرداختی و مانده)  ·  "
                      "در باسکول: «خرید شراکتی (ورود)» و «فروش شراکتی (خروج)»")
    L.st(c, name=FA, size=10.5, bold=True, color=L.WHITE, bg=L.TERRA, border=L.TERRA, h="right")
    ws.row_dimensions[r].height = 24
    cell = ws.cell(r + 1, 1, "حساب شرکا  Shrk")
    cell.hyperlink = "#'Shrk'!A1"
    L.st(cell, name=FA, size=10, bold=True, color=L.FOREST, bg=L.MINT, border=L.LINE)
    ws.print_area = f"A1:J{r + 1}"
    return r


def entry(path, master_path):
    partners = read_partners(master_path)
    wb = openpyxl.load_workbook(path)
    _sync(wb, partners)
    n_ops = _sync_ops(wb, master_path)
    n_demo = demo_rows(wb)
    sc = wb["Scale"]
    n_cols = scale_columns(sc)
    n_legs = scale_legs(sc)
    hide = scale_hide_helpers(sc)
    for col, text in SCALE_TIPS.items():
        L.tip(sc, f"{col}8", text, w=320, h=130)
    # partner codes stay visible in the payment sheet tooltip
    pay = wb["Pay"]
    L.tip(pay, "F8", "کد طرف حساب: باغدار C، بارفروش B/SL، باسکول دیگران BD، شریک SHxx. "
                    "برای پرداخت سهم شراکت، کد شریک را انتخاب کنید.", w=320, h=120)
    report(wb, partners)
    entry_help(wb)
    entry_home(wb)
    sc.print_area = f"A1:T{grow.SCALE_LAST}"
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)
    wb.close()
    return {"partners": len([p for p in partners if p[1]]), "ops": n_ops,
            "demo": n_demo, "cols": n_cols, "legs": n_legs, "hidden": hide}

# =================================================================== mirrors
ACC_NAMES = ("AcctTbl", "AcctCodes", "ListAllCodes", "ListAcctNames")


def grow_account_ranges(wb, last=SYNC_LAST):
    """Account tables 10…259 → 10…279 (a superset: no formula changes)."""
    n = 0
    for name, dn in list(wb.defined_names.items()):
        if name not in ACC_NAMES:
            continue
        txt = dn.attr_text
        new = re.sub(r"\$([A-Z]{1,3})\$10:\$([A-Z]{1,3})\$259",
                     lambda m: f"${m.group(1)}$10:${m.group(2)}${last}", txt)
        new = new.replace("$C$259", f"$C${last}")
        if new != txt:
            dn.attr_text = new
            n += 1
    return n


def renumber(ws, last, col="B"):
    """Column A = consecutive row number of every filled row."""
    for r in range(10, last + 1):
        ws[f"A{r}"].value = f'=IF(${col}{r}="","",COUNTA(${col}$10:${col}{r}))'
    return last - 9
