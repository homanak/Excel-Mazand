# -*- coding: utf-8 -*-
"""02_Bank.xlsx — the transaction file (v03).

Entry sheet + journal + 5 analysis sheets + cheque register + sync + help.
Every dropdown item is defined in the master file (00_Master.xlsx) and copied
into the locked Sync sheet of this file.
"""
import openpyxl
import xlsxwriter

from kit import (Kit, FOREST, FOREST_2, FOREST_3, SAGE, SAGE_L, MINT, CREAM, SAND, SAND_2,
                 GOLD, GOLD_L, AMBER, TERRA, BLUSH, SKY, SKY_D, LAVENDER, ROSE, PEACH,
                 GREY_BG, LINE, LINE_SOFT, INK, MUTED, WHITE, RED, GREEN, BLUE,
                 FA, FA_TITLE, EN, INT, INT_RED, DEC, PCT, COMPANY, COMPANY_EN)

FILE = "02_Bank.xlsx"
NR = 140          # transaction rows  (12..151)
PN = NR * 2       # journal lines      (12..291)

BANK_OPS = [
    # title, direction, party mode, fixed party, note
    ("دریافت از مشتری", "IN", "CODE", "", "واریز مشتری به حساب شرکت — طرف حساب را با کد وارد کنید"),
    ("پرداخت به مشتری", "OUT", "CODE", "", "پرداخت/حواله شرکت به طرف حساب"),
    ("دریافت نقدی و واریز به بانک", "IN", "FIXED", "X003", "پول نقد از صندوق به بانک می‌رود"),
    ("پرداخت هزینه", "OUT", "CODE", "", "هزینه‌ها: حمل، دستمزد، کارمزد، مالیات…"),
    ("انتقال داخلی بین بانک‌ها", "TRF", "NONE", "", "فقط جابه‌جایی بین بانک‌های خود شرکت؛ تراز شرکت تغییر نمی‌کند"),
    ("پرداخت انعام مصطفی", "OUT", "FIXED", "X001", "بدهی انعام مصطفی تسویه می‌شود"),
    ("دریافت هزینه باسکول دیگران", "IN", "PICK", "", "پول از طرف باسکول دیگران؛ در حساب همان شخص بستانکار می‌شود"),
    ("پرداخت به باسکول دیگران", "OUT", "PICK", "", "تسویه با شخص باسکول دیگران"),
    ("واریز پز باسکول (۲۰۰۶)", "IN", "FIXED", "X011", "سرجمع پز روز قبل؛ حساب واسط X011 را تسویه می‌کند (درآمد در دفتر روزنامه ثبت شده)"),
    ("پاس چک (پرداخت چک)", "OUT", "FIXED", "X010", "چک پرداختی پاس می‌شود"),
    ("برگشت چک دریافتی", "IN", "FIXED", "X010", "چک دریافتی برگشت خورد"),
    ("واریز چک به حساب", "IN", "FIXED", "X010", "چک دریافتی وصول شد"),
    ("دریافت تسهیلات", "IN", "CODE", "", "وام/تسهیلات بانکی"),
    ("بازپرداخت تسهیلات", "OUT", "CODE", "", "قسط یا اصل وام"),
    ("سایر دریافت", "IN", "CODE", "", ""),
    ("سایر پرداخت", "OUT", "CODE", "", ""),
    ("تسویه دریافت (واسط)", "IN", "FIXED", "X011", "واریزی‌ای که قبلاً در شیت پرداخت (Pay) ثبت شده؛ واسط X011 را صفر می‌کند"),
    ("تسویه پرداخت (واسط)", "OUT", "FIXED", "X011", "پرداختی که قبلاً در شیت پرداخت (Pay) ثبت شده؛ واسط X011 را صفر می‌کند"),
]

DEMO = [
    ("1405/07/01", "TB-001", "واریز پز باسکول (۲۰۰۶)", "BN02", "", "کارت/پوز", "", "واریز سرجمع پز ۱۴۰۵/۰۶/۳۱", "", 120000000),
    ("1405/07/01", "TB-002", "دریافت از مشتری", "BN01", "", "انتقال بانکی", "C001", "علی‌الحساب خرید مرکبات", "TR-8841", 150000000),
    ("1405/07/02", "TB-003", "پرداخت انعام مصطفی", "BN02", "", "کارت/پوز", "", "انعام هفته اول", "", 5000000),
    ("1405/07/02", "TB-004", "انتقال داخلی بین بانک‌ها", "BN01", "BN03", "انتقال بانکی", "", "تقویت حساب ملت", "TR-8842", 200000000),
    ("1405/07/03", "TB-005", "پرداخت به مشتری", "BN03", "", "حواله", "C002", "تسویه بخشی از حساب باغدار", "TR-8843", 60000000),
    ("1405/07/03", "TB-006", "دریافت هزینه باسکول دیگران", "BN02", "", "کارت/پوز", "", "هزینه باسکول میدان بار", "", 1500000, "BD001 – باسکول میدان بار (نمونه)"),
    ("1405/07/04", "TB-007", "واریز پز باسکول (۲۰۰۶)", "BN02", "", "کارت/پوز", "", "واریز پز ۱۴۰۵/۰۷/۰۳", "", 95000000),
    ("1405/07/05", "TB-008", "پرداخت هزینه", "BN01", "", "انتقال بانکی", "X012", "هزینه حمل و تخلیه", "", 12000000),
    ("1405/07/06", "TB-009", "پاس چک (پرداخت چک)", "BN01", "", "پاس چک", "", "چک ۴۵۲۱ به تأمین‌کننده", "4521", 40000000),
    ("1405/07/07", "TB-010", "دریافت از مشتری", "BN03", "", "انتقال بانکی", "C003", "واریز بابت فاکتور بار", "TR-8844", 88000000),
    ("1405/07/08", "TB-011", "دریافت نقدی و واریز به بانک", "BN01", "", "نقد", "", "واریز نقدی صندوق", "", 25000000),
    ("1405/07/09", "TB-012", "پرداخت به مشتری", "BN02", "", "کارت/پوز", "S001", "پرداخت به فروشنده سبد", "", 33000000),
    ("1405/07/10", "TB-013", "واریز پز باسکول (۲۰۰۶)", "BN02", "", "کارت/پوز", "", "واریز پز ۱۴۰۵/۰۷/۰۹", "", 76000000),
    ("1405/07/11", "TB-014", "پرداخت به باسکول دیگران", "BN02", "", "کارت/پوز", "", "تسویه با باسکول میدان بار", "", 1000000, "BD001 – باسکول میدان بار (نمونه)"),
    ("1405/07/12", "TB-015", "سایر دریافت", "BN03", "", "انتقال بانکی", "C005", "دریافت متفرقه", "", 9500000),
]

CHEQUES = [
    ("1405/07/02", "1405/08/02", "4521", "پرداختی", "X012", "BN01", 40000000, "پاس شده", "1405/07/06", "حمل و تخلیه"),
    ("1405/07/04", "1405/08/04", "7732", "دریافتی", "C002", "BN03", 65000000, "در جریان", "", "چک باغدار"),
    ("1405/07/06", "1405/09/06", "8890", "پرداختی", "X013", "BN01", 18000000, "در جریان", "", "دستمزد"),
    ("1405/07/08", "1405/08/08", "1120", "دریافتی", "SL001", "BN02", 42000000, "برگشتی", "", "برگشت خورد"),
    ("1405/07/10", "1405/09/10", "9031", "پرداختی", "X012", "BN03", 22000000, "باطل", "", "باطل شد"),
    ("1405/07/11", "1405/08/11", "6654", "دریافتی", "C004", "BN01", 51000000, "در جریان", "", ""),
]


def master_data(path):
    wb = openpyxl.load_workbook(path)
    acc = []
    ws = wb["Acc"]
    for r in range(10, 260):
        code = ws.cell(r, 1).value
        if not code:
            continue
        acc.append((str(code), ws.cell(r, 4).value or "", ws.cell(r, 2).value or "",
                    ws.cell(r, 3).value or "", ws.cell(r, 9).value or "", ws.cell(r, 10).value or ""))
    banks = []
    for r in range(10, 20):
        b = wb["Bnk"]
        if not b.cell(r, 1).value:
            continue
        banks.append((b.cell(r, 1).value, b.cell(r, 2).value, b.cell(r, 3).value or 0,
                      b.cell(r, 4).value, b.cell(r, 6).value or "بله", b.cell(r, 7).value or ""))
    groups = [wb["Grp"].cell(r, 1).value for r in range(10, 55) if wb["Grp"].cell(r, 1).value]
    wb.close()
    seen, gl = set(), []
    for g in groups:
        if g not in seen:
            seen.add(g)
            gl.append(g)
    bd = [a for a in acc if a[0].startswith("BD")]
    return acc, banks, gl, bd


# =====================================================================  Sync
def sheet_sync(wb, k, acc, banks, gl, bd):
    ws = wb.add_worksheet("Sync")
    k.setup(ws, 12, 262, landscape=True, repeat=(8, 9), freeze=(9, 3))
    ws.set_column(0, 0, 11)
    ws.set_column(1, 1, 34)
    ws.set_column(2, 2, 15)
    ws.set_column(3, 3, 13)
    ws.set_column(4, 4, 8)
    ws.set_column(5, 5, 34)
    ws.set_column(6, 6, 2)
    ws.set_column(7, 7, 9)
    ws.set_column(8, 8, 8)
    ws.set_column(9, 9, 26)
    ws.set_column(10, 10, 16)
    ws.set_column(11, 11, 15)
    ws.set_column(12, 12, 12)
    ws.set_column(13, 13, 2)
    ws.set_column(14, 14, 26)
    ws.set_column(15, 15, 8)
    ws.set_column(16, 16, 9)
    ws.set_column(17, 17, 9)
    ws.set_column(18, 18, 26)
    ws.set_column(19, 19, 9)
    ws.set_column(20, 20, 8)
    ws.set_column(21, 21, 22)
    ws.set_column(22, 22, 34)

    ws.merge_range(0, 0, 0, 22, "داده‌های ثابت — کپی قفل‌شده از فایل مادر 00_Master.xlsx",
                   k.f(font_name=FA_TITLE, font_size=15, bold=True, font_color=FOREST,
                       bg_color=SAND_2, align="right", valign="vcenter"))
    ws.merge_range(1, 0, 1, 22,
                   "این شیت فقط مطالعه است. هر کشوی این فایل از همین‌جا می‌خواند؛ برای تغییر، فایل مادر را ویرایش کنید و این فایل را دوباره بسازید.",
                   k.f(font_name=FA, font_size=9.5, font_color=MUTED, align="right",
                       valign="vcenter", text_wrap=True))
    ws.set_row(1, 22)

    def head(r, c0, c1, fa, en, tint=FOREST):
        if c1 > c0:
            ws.merge_range(r, c0, r, c1, fa, k.f(font_name=FA, font_size=9.5, bold=True,
                                                 font_color=WHITE, bg_color=tint,
                                                 align="center", valign="vcenter", border=1,
                                                 border_color=tint))
        else:
            ws.write(r, c0, fa, k.f(font_name=FA, font_size=9.5, bold=True, font_color=WHITE,
                                    bg_color=tint, align="center", valign="vcenter",
                                    border=1, border_color=tint))
        ws.write(r + 1, c0, en, k.f(font_name=EN, font_size=6.5, font_color=SAGE_L,
                                    bg_color=tint, align="center", valign="vcenter"))

    f_c = k.f(font_name=EN, font_size=8.5, bg_color=WHITE, border=1, border_color=LINE_SOFT,
              align="center", valign="vcenter")
    f_t = k.f(font_name=FA, font_size=9, bg_color=WHITE, border=1, border_color=LINE_SOFT,
              align="right", valign="vcenter")
    f_n = k.f(font_name=EN, font_size=8.5, bg_color=WHITE, border=1, border_color=LINE_SOFT,
              align="center", valign="vcenter", num_format=INT)

    # ---- banks  A..F  rows 10..19
    head(8, 0, 5, "بانک‌ها و حساب‌های شرکت", "BANKS", FOREST)
    for i, (code, title, opening, accode, active, note) in enumerate(banks):
        r = 9 + i
        ws.write(r, 0, code, f_c)
        ws.write(r, 1, title, f_t)
        ws.write_number(r, 2, opening, f_n)
        ws.write(r, 3, accode, f_c)
        ws.write(r, 4, active, f_c)
        ws.write(r, 5, note, f_t)

    # ---- accounts H..L rows 10..259
    head(8, 7, 11, "حساب‌ها (همه طرف‌حساب‌ها و حساب‌های معین)", "ACCOUNTS", TERRA)
    for i, (code, name, grp, sub, typ, act) in enumerate(acc):
        r = 9 + i
        ws.write(r, 7, code, f_c)
        ws.write(r, 8, name, f_t)
        ws.write(r, 9, grp, f_t)
        ws.write(r, 10, sub, f_t)
        ws.write(r, 11, typ, f_c)

    # ---- bank operations N..S rows 10..25
    head(8, 13, 18, "نوع عملیات بانکی — قاعده بدهکار/بستانکار", "BANK OPERATIONS", FOREST_2)
    for i, (title, direction, mode, fixed, note) in enumerate(BANK_OPS):
        r = 9 + i
        ws.write(r, 13, title, f_t)
        ws.write(r, 14, direction, f_c)
        ws.write(r, 15, mode, f_c)
        ws.write(r, 16, fixed, f_c)
        ws.write(r, 17, f"B{i+1:02d}", f_c)
        ws.write(r, 18, note, k.f(font_name=FA, font_size=8, bg_color=WHITE,
                                 border=1, border_color=LINE_SOFT, align="right",
                                 valign="vcenter", text_wrap=True))

    # ---- lists at row 8-9 starting col 19: pay methods, cheque status, BD labels
    head(8, 19, 19, "روش پرداخت", "METHODS", GOLD)
    for i, m in enumerate(["نقد", "چک", "انتقال بانکی", "کارت/پوز", "حواله"]):
        ws.write(9 + i, 19, m, f_t)
    head(8, 20, 20, "وضعیت چک", "CHEQUE", GOLD)
    for i, m in enumerate(["در جریان", "پاس شده", "برگشتی", "باطل"]):
        ws.write(9 + i, 20, m, f_t)
    head(8, 21, 21, "شخص باسکول دیگران", "OTHER SCALES", GOLD)
    labels = [f"{c} – {n}" if n else f"{c} – (بدون نام)" for c, n, *_ in bd]
    for i, m in enumerate(labels):
        ws.write(9 + i, 21, m, f_t)
    head(8, 23, 23, "نوع چک", "CHEQUE TYPE", GOLD)
    for i, m in enumerate(["دریافتی", "پرداختی"]):
        ws.write(9 + i, 23, m, f_t)
    head(8, 22, 22, "گروه‌های حساب", "GROUPS", GOLD)
    for i, g in enumerate(gl[:24]):
        ws.write(9 + i, 22, g, f_t)

    # ---- names
    wb.define_name("BankTbl", "=Sync!$A$10:$F$19")
    wb.define_name("BankAccts", "=Sync!$D$10:$D$19")
    wb.define_name("ListBanks", "=Sync!$A$10:$A$19")
    wb.define_name("AcctTbl", "=Sync!$H$10:$L$259")
    wb.define_name("ListAllCodes", "=Sync!$H$10:$H$259")
    wb.define_name("BankOpTbl", "=Sync!$N$10:$S$27")
    wb.define_name("ListBankOps", "=Sync!$N$10:$N$27")
    wb.define_name("ListPayMethods", "=Sync!$T$10:$T$14")
    wb.define_name("ListChequeStatus", "=Sync!$U$10:$U$13")
    wb.define_name("ListOtherScales", "=Sync!$V$10:$V$29")
    wb.define_name("ListGroups", "=Sync!$W$10:$W$33")
    wb.define_name("ListChequeTypes", "=Sync!$X$10:$X$11")

    ws.protect("", {"select_locked_cells": True, "select_unlocked_cells": False,
                    "format_cells": False, "format_columns": False, "format_rows": False})
    return ws


# =====================================================================  Txn
COLS = [
    ("ردیف", "No.", 6, "شماره خودکار ردیف؛ با پر شدن تاریخ ساخته می‌شود."),
    ("تاریخ", "Date", 11, "تاریخ شمسی به شکل 1405/07/10. فقط عدد وارد کنید؛ اسلش‌ها را خودتان بزنید."),
    ("شماره سند", "Doc", 11, "شماره پیگیری یا شماره سند داخلی؛ آزاد."),
    ("نوع عملیات", "Operation", 24, "از کشو انتخاب کنید. جهت بدهکار/بستانکار و حساب طرف از همین انتخاب تعیین می‌شود."),
    ("بانک عامل", "Company bank", 20, "بانک شرکت که پول از آن خارج یا به آن وارد می‌شود (برای انتقال داخلی: بانک مبدأ)."),
    ("بانک مقصد", "Destination", 18, "فقط در «انتقال داخلی» پر می‌شود: بانکی که پول به آن واریز می‌شود."),
    ("روش", "Method", 13, "نقد، چک، انتقال بانکی، کارت/پوز یا حواله."),
    ("کد طرف حساب", "Party code", 13, "کد را تایپ کنید (مثل C001 یا X012)؛ نام و گروه خودکار می‌آید. برای باسکول دیگران این ستون را خالی بگذارید."),
    ("نام طرف حساب", "Party name", 22, "خودکار از فایل مادر."),
    ("گروه", "Group", 14, "خودکار؛ برای گزارش‌های تحلیلی."),
    ("شرح", "Description", 26, "توضیح کوتاه؛ در گزارش‌ها دیده می‌شود."),
    ("شماره پیگیری / چک", "Ref / Cheque", 14, "شماره پیگیری حواله، چک یا فیش."),
    ("مبلغ (ریال)", "Amount", 15, "مبلغ کامل تراکنش؛ بدهکار و بستانکار خودکار از آن ساخته می‌شود."),
    ("شخص باسکول دیگران", "Other scale", 22, "فقط برای عملیات باسکول دیگران: شخص را از کشو انتخاب کنید تا بدهی‌اش قابل ردیابی شود."),
    ("بدهکار بانک", "Bank Dr", 14, "خودکار: پول وارد حساب بانک شده است."),
    ("بستانکار بانک", "Bank Cr", 14, "خودکار: پول از حساب بانک خارج شده است."),
    ("بدهکار طرف", "Party Dr", 13, "خودکار: طرف حساب/بانک مقصد بدهکار می‌شود."),
    ("بستانکار طرف", "Party Cr", 13, "خودکار: طرف حساب بستانکار می‌شود."),
    ("وضعیت", "Status", 12, "✓ یعنی ردیف کامل و قابل ثبت است؛ ⚠ یعنی چیزی جا افتاده."),
    ("یادداشت", "Note", 22, "هر توضیح اضافه."),
]


def sheet_txn(wb, k):
    ws = wb.add_worksheet("Txn")
    last = len(COLS) - 1
    k.setup(ws, last, 12 + NR + 2, landscape=True, repeat=(11, 11), freeze=(12, 4))
    k.brand(ws, last, "تراکنش‌های بانکی — ثبت، کنترل و ردیابی",
            "بانک عامل را از کشو انتخاب کنید  ·  کد طرف حساب را تایپ کنید  ·  انتقال داخلی: بانک مبدأ و مقصد",
            icon="bank")
    k.kpis(ws, 6, [
        ("تعداد تراکنش", f'=COUNTA($B$13:$B${12+NR})', INT, FOREST, "list"),
        ("جمع دریافت", f'=SUM($O$13:$O${12+NR})', INT, GREEN, "transfer"),
        ("جمع پرداخت", f'=SUM($P$13:$P${12+NR})', INT, RED, "transfer"),
        ("مانده جریان", f'=SUM($O$13:$O${12+NR})-SUM($P$13:$P${12+NR})', INT_RED, BLUE, "money"),
        ("ردیف‌های ناقص", f'=COUNTIF($S$13:$S${12+NR},"⚠*")', INT, TERRA, "warn"),
    ], last)

    k.note(ws, 8, last,
           "بانک عامل همیشه از فهرست بانک‌های شرکت انتخاب می‌شود. ستون‌های خاکستری خودکارند و دست نزنید. "
           "برای انتقال داخلی، بانک مقصد را هم انتخاب کنید تا پول از یک حساب کم و به حساب دیگر اضافه شود.",
           bg=SAND, icon="flag")

    k.section(ws, 10, last, "۰۱", "دفتر تراکنش‌ها", "هر ردیف یک تراکنش بانکی")
    k.thead(ws, 11, COLS, FOREST)

    f_in = k.f(font_name=EN, font_size=9, bg_color="#FFF4D6", border=1,
               border_color=GOLD_L, align="center", valign="vcenter")
    f_in_t = k.f(font_name=FA, font_size=9.5, bg_color="#FFF4D6", border=1,
                 border_color=GOLD_L, align="right", valign="vcenter")
    f_auto = k.f(font_name=EN, font_size=8.5, bg_color="#F6F5F0", border=1,
                 border_color=LINE_SOFT, align="center", valign="vcenter")
    f_auto_t = k.f(font_name=FA, font_size=9, bg_color="#F6F5F0", border=1,
                   border_color=LINE_SOFT, align="right", valign="vcenter")
    f_auto_n = k.f(font_name=EN, font_size=9, bg_color="#F6F5F0", border=1,
                   border_color=LINE_SOFT, align="center", valign="vcenter", num_format=INT)
    f_status = k.f(font_name=FA, font_size=9, bold=True, bg_color=CREAM, border=1,
                   border_color=LINE_SOFT, align="center", valign="vcenter")

    for i in range(NR):
        r = 12 + i
        x = r + 1
        ws.set_row(r, 21)
        ws.write_formula(r, 0, f'=IF($B{x}="","",COUNTA($B$13:$B{x}))', f_auto)
        ws.write_blank(r, 1, None, f_in)
        ws.write_blank(r, 2, None, f_in)
        ws.write_blank(r, 3, None, f_in_t)
        ws.write_blank(r, 4, None, f_in_t)
        ws.write_blank(r, 5, None, f_in_t)
        ws.write_blank(r, 6, None, f_in)
        ws.write_blank(r, 7, None, f_in)
        ws.write_formula(r, 8, f'=IF($H{x}="","",IFERROR(VLOOKUP($H{x},AcctTbl,2,0),"کد ناشناخته"))', f_auto_t)
        ws.write_formula(r, 9, f'=IF($H{x}="","",IFERROR(VLOOKUP($H{x},AcctTbl,3,0),""))', f_auto)
        ws.write_blank(r, 10, None, f_in_t)
        ws.write_blank(r, 11, None, f_in)
        ws.write_blank(r, 12, None, k.f(font_name=EN, font_size=9, bg_color="#FFF4D6",
                                        border=1, border_color=GOLD_L, align="center",
                                        valign="vcenter", num_format=INT))
        ws.write_blank(r, 13, None, f_in_t)
        ws.write_formula(r, 14, f'=IF($B{x}="",0,IF($U{x}="IN",N($M{x}),0))', f_auto_n)
        ws.write_formula(r, 15, f'=IF($B{x}="",0,IF(OR($U{x}="OUT",$U{x}="TRF"),N($M{x}),0))', f_auto_n)
        ws.write_formula(r, 16, f'=IF($B{x}="",0,IF(OR($U{x}="OUT",$U{x}="TRF"),N($M{x}),0))', f_auto_n)
        ws.write_formula(r, 17, f'=IF($B{x}="",0,IF($U{x}="IN",N($M{x}),0))', f_auto_n)
        ws.write_formula(
            r, 18,
            f'=IF($B{x}="","",IF($M{x}="","⚠ مبلغ",IF($E{x}="","⚠ بانک عامل",'
            f'IF(AND($U{x}="TRF",$F{x}=""),"⚠ بانک مقصد",'
            f'IF(AND($V{x}<>"NONE",$V{x}<>"",$Z{x}=""),"⚠ طرف حساب","✓")))))', f_status)
        ws.write_blank(r, 19, None, f_in_t)
        # hidden helpers
        ws.write_formula(r, 20, f'=IFERROR(VLOOKUP($D{x},BankOpTbl,2,0),"")', f_auto)
        ws.write_formula(r, 21, f'=IFERROR(VLOOKUP($D{x},BankOpTbl,3,0),"")', f_auto)
        ws.write_formula(r, 22, f'=IFERROR(VLOOKUP($D{x},BankOpTbl,4,0),"")', f_auto)
        ws.write_formula(r, 23, f'=IF($E{x}="","",IFERROR(VLOOKUP($E{x},BankTbl,4,0),""))', f_auto)
        ws.write_formula(r, 24, f'=IF($F{x}="","",IFERROR(VLOOKUP($F{x},BankTbl,4,0),""))', f_auto)
        ws.write_formula(
            r, 25,
            f'=IF($B{x}="","",IF($V{x}="PICK",LEFT($N{x},5),IF($V{x}="FIXED",$W{x},'
            f'IF($V{x}="CODE",$H{x},""))))', f_auto)

    for c, w in ((20, 7), (21, 8), (22, 9), (23, 10), (24, 10), (25, 11)):
        ws.set_column(c, c, w)
        ws.set_column(c, c, w, None, {"hidden": True})

    # demo rows
    for i, row in enumerate(DEMO):
        r = 12 + i
        date, doc, op, bank, dest, method, code, desc, ref, amount = row[:10]
        bd = row[10] if len(row) > 10 else ""
        ws.write_string(r, 1, date, f_in)
        ws.write_string(r, 2, doc, f_in)
        ws.write_string(r, 3, op, f_in_t)
        ws.write_string(r, 4, bank, f_in_t)
        if dest:
            ws.write_string(r, 5, dest, f_in_t)
        ws.write_string(r, 6, method, f_in)
        if code:
            ws.write_string(r, 7, code, f_in)
        ws.write_string(r, 10, desc, f_in_t)
        if ref:
            ws.write_string(r, 11, ref, f_in)
        ws.write_number(r, 12, amount, k.f(font_name=EN, font_size=9, bg_color="#FFF4D6",
                                           border=1, border_color=GOLD_L, align="center",
                                           valign="vcenter", num_format=INT))
        if bd:
            ws.write_string(r, 13, bd, f_in_t)

    total = 12 + NR
    ws.merge_range(total, 0, total, 13, "جمع کل",
                   k.f(font_name=FA, font_size=10.5, bold=True, font_color=WHITE,
                       bg_color=GOLD, align="center", valign="vcenter", border=1,
                       border_color=GOLD))
    for c in (14, 15, 16, 17):
        ws.write_formula(total, c, f"=SUM({chr(65+c)}12:{chr(65+c)}{total-1})",
                         k.f(font_name=EN, font_size=10, bold=True, font_color=WHITE,
                            bg_color=GOLD, align="center", valign="vcenter", border=1,
                            border_color=GOLD, num_format=INT))
    ws.write_formula(total, 18, f'=IF(COUNTIF($S$13:$S${total},"⚠*")=0,"✓ تراز","⚠ بررسی")',
                     k.f(font_name=FA, font_size=10, bold=True, font_color=WHITE,
                        bg_color=GOLD, align="center", valign="vcenter", border=1,
                        border_color=GOLD))
    ws.set_row(total, 24)

    # validations
    ws.data_validation(12, 1, 12 + NR - 1, 1, {
        "validate": "custom",
        "value": '=OR($B13="",AND(LEN($B13)=10,MID($B13,5,1)="/",MID($B13,8,1)="/"))',
        "error_title": "تاریخ شمسی",
        "error_message": "تاریخ را به شکل 1405/07/10 وارد کنید.",
        "error_type": "stop"})
    ws.data_validation(12, 3, 12 + NR - 1, 3, {"validate": "list", "source": "=ListBankOps",
                                               "error_type": "stop"})
    ws.data_validation(12, 4, 12 + NR - 1, 4, {"validate": "list", "source": "=ListBanks",
                                               "error_type": "stop"})
    ws.data_validation(12, 5, 12 + NR - 1, 5, {"validate": "list", "source": "=ListBanks",
                                               "error_type": "warning"})
    ws.data_validation(12, 6, 12 + NR - 1, 6, {"validate": "list", "source": "=ListPayMethods",
                                               "error_type": "warning"})
    ws.data_validation(12, 7, 12 + NR - 1, 7, {"validate": "list", "source": "=ListAllCodes",
                                               "error_type": "warning"})
    ws.data_validation(12, 13, 12 + NR - 1, 13, {"validate": "list", "source": "=ListOtherScales",
                                                 "error_type": "warning"})

    # colour by direction
    ws.conditional_format(11, 0, 12 + NR - 1, last, {
        "type": "formula", "criteria": '=$U12="IN"',
        "format": k.f(bg_color=MINT)})
    ws.conditional_format(11, 0, 12 + NR - 1, last, {
        "type": "formula", "criteria": '=$U12="OUT"',
        "format": k.f(bg_color=BLUSH)})
    ws.conditional_format(11, 0, 12 + NR - 1, last, {
        "type": "formula", "criteria": '=$U12="TRF"',
        "format": k.f(bg_color=LAVENDER)})
    ws.conditional_format(11, 18, 12 + NR - 1, 18, {
        "type": "text", "criteria": "containing", "value": "⚠",
        "format": k.f(bg_color="#FBE3D8", font_color=TERRA, bold=True)})
    ws.conditional_format(11, 18, 12 + NR - 1, 18, {
        "type": "text", "criteria": "containing", "value": "✓",
        "format": k.f(bg_color=MINT, font_color=GREEN, bold=True)})
    ws.autofilter(10, 0, 12 + NR - 1, last)
    return ws


# =====================================================================  Post
def sheet_post(wb, k):
    ws = wb.add_worksheet("Post")
    ws.hide()
    k.setup(ws, 13, 292, landscape=True, freeze=(12, 1))
    cols = ["تاریخ", "سند", "منبع", "کد حساب", "نام حساب", "گروه", "عملیات", "بانک",
            "بدهکار", "بستانکار", "ماه", "شرح", "", "بانک؟"]
    for i, h in enumerate(cols):
        ws.write(8, i, h, k.f(font_name=FA, font_size=8, bold=True, font_color=WHITE,
                              bg_color=FOREST, align="center"))
    f_t = k.f(font_name=EN, font_size=8, align="right")
    f_n = k.f(font_name=EN, font_size=8, num_format=INT, align="center", bg_color="#F6F5F0")
    f_free = k.f(font_name=FA, font_size=8, align="right", bg_color="#F6F5F0")
    for i in range(NR):
        x = 13 + i
        p = 12 + 2 * i
        for j, (row, dr, cr, acct, desc) in enumerate((
                (p, "O", "P", f"Txn!$X{x}",
                 f'IF(Txn!$K{x}="","",Txn!$K{x})'),
                (p + 1, "Q", "R",
                 f'IF(Txn!$U{x}="TRF",Txn!$Y{x},Txn!$Z{x})',
                 f'IF(Txn!$K{x}="","","طرف مقابل — "&Txn!$K{x})'))):
            ws.write_formula(row, 0, f'=IF(Txn!$B{x}="","",Txn!$B{x})', f_t)
            ws.write_formula(row, 1, f'=IF(Txn!$B{x}="","",Txn!$C{x})', f_t)
            ws.write_formula(row, 2, f'=IF(Txn!$B{x}="","","تراکنش بانکی")', f_free)
            ws.write_formula(row, 3, f'=IF(Txn!$B{x}="","",IF({acct}="","",{acct}))', f_t)
            ws.write_formula(row, 4, f'=IF($D{row+1}="","",IFERROR(VLOOKUP($D{row+1},AcctTbl,2,0),""))', f_free)
            ws.write_formula(row, 5, f'=IF($D{row+1}="","",IFERROR(VLOOKUP($D{row+1},AcctTbl,3,0),""))', f_free)
            ws.write_formula(row, 6, f'=IF(Txn!$B{x}="","",Txn!$D{x})', f_free)
            ws.write_formula(row, 7, f'=IF(Txn!$B{x}="","",Txn!$E{x})', f_t)
            ws.write_formula(row, 8, f'=IF(Txn!$B{x}="",0,Txn!${dr}{x})', f_n)
            ws.write_formula(row, 9, f'=IF(Txn!$B{x}="",0,Txn!${cr}{x})', f_n)
            ws.write_formula(row, 10,
                             f'=IF($A{row+1}="","",IFERROR(VALUE(MID($A{row+1},6,2)),""))',
                             k.f(font_name=EN, font_size=8, num_format="00", align="center",
                                 bg_color="#F6F5F0"))
            ws.write_formula(row, 11, f'={desc}', f_free)
            ws.write_formula(row, 13,
                             f'=IF($D{row+1}="",0,COUNTIF(BankAccts,$D{row+1}))',
                             k.f(font_name=EN, font_size=8, align="center",
                                 bg_color="#F6F5F0"))
    for c, w in enumerate([11, 11, 14, 11, 26, 14, 24, 11, 14, 14, 7, 28]):
        ws.set_column(c, c, w)
    return ws


# =====================================================================  Sum
def sheet_sum(wb, k, banks):
    ws = wb.add_worksheet("Sum")
    last = 7
    k.setup(ws, last, 26, landscape=True, repeat=(10, 10), freeze=(11, 2))
    k.brand(ws, last, "مانده و گردش حساب‌های بانکی",
            "مانده هر بانک = مانده اولیه + دریافت‌ها − پرداخت‌ها  ·  همه از دفتر تراکنش‌ها", icon="bank")
    k.kpis(ws, 6, [
        ("جمع مانده بانک‌ها", "=SUM($F$12:$F$21)", INT, FOREST, "money"),
        ("جمع دریافت", "=SUM($C$12:$C$21)", INT, GREEN, "transfer"),
        ("جمع پرداخت", "=SUM($D$12:$D$21)", INT, RED, "transfer"),
        ("خالص جریان", "=$C$22-$D$22", INT_RED, BLUE, "chart"),
    ], last)

    k.section(ws, 9, last, "۰۱", "وضعیت هر بانک", "به ترتیب مانده")
    heads = [("بانک", "Bank", 26), ("مانده اولیه", "Opening", 15),
             ("دریافت", "Inflow", 15), ("پرداخت", "Outflow", 15),
             ("مانده فعلی", "Balance", 16), ("سهم", "Share", 9),
             ("شماره حساب", "Acct", 11), ("یادداشت", "Note", 30)]
    for i, (fa, en, w) in enumerate(heads):
        ws.set_column(i, i, w)
        ws.write(10, i, fa, k.f(font_name=FA, font_size=9.5, bold=True, font_color=WHITE,
                                bg_color=FOREST, border=1, border_color=FOREST,
                                align="center", valign="vcenter"))
    for i, (code, title, opening, accode, active, note) in enumerate(banks):
        r = 11 + i
        ws.set_row(r, 22)
        ws.write_formula(r, 0, f'=BankTbl[{i+1},2]' if False else f'=INDEX(BankTbl,{i+1},2)',
                         k.f(font_name=FA, font_size=9.5, bold=True, font_color=FOREST,
                             bg_color=SAND_2, border=1, border_color=LINE, align="right",
                             valign="vcenter"))
        ws.write_formula(r, 1, f'=INDEX(BankTbl,{i+1},3)',
                         k.f(font_name=EN, font_size=9, bg_color=WHITE, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(r, 2, f'=SUMIFS(Post!$I$13:$I${12+PN},Post!$D$13:$D${12+PN},$G{r+1})',
                         k.f(font_name=EN, font_size=9, bg_color=WHITE, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(r, 3, f'=SUMIFS(Post!$J$13:$J${12+PN},Post!$D$13:$D${12+PN},$G{r+1})',
                         k.f(font_name=EN, font_size=9, bg_color=WHITE, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(r, 4, f'=$B{r+1}+$C{r+1}-$D{r+1}',
                         k.f(font_name=EN, font_size=10, bold=True, bg_color=CREAM, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT_RED))
        ws.write_formula(r, 5, f'=IFERROR($E{r+1}/SUM($E$12:$E$21),0)',
                         k.f(font_name=EN, font_size=9, bg_color=WHITE, border=1,
                             border_color=LINE_SOFT, align="center", num_format=PCT))
        ws.write_formula(r, 6, f'=INDEX(BankTbl,{i+1},4)',
                         k.f(font_name=EN, font_size=8.5, bg_color="#F6F5F0", border=1,
                             border_color=LINE_SOFT, align="center"))
        ws.write_formula(r, 7, f'=INDEX(BankTbl,{i+1},6)',
                         k.f(font_name=FA, font_size=8.5, bg_color="#F6F5F0", border=1,
                             border_color=LINE_SOFT, align="right"))
    ws.set_row(21, 24)
    ws.write(21, 0, "جمع",
             k.f(font_name=FA, font_size=10, bold=True, font_color=WHITE,
                 bg_color=GOLD, align="center", border=1, border_color=GOLD))
    for c in (1, 2, 3, 4):
        L = chr(65 + c)
        ws.write_formula(21, c, f"=SUM({L}12:{L}21)",
                         k.f(font_name=EN, font_size=10, bold=True, font_color=WHITE,
                            bg_color=GOLD, align="center", border=1, border_color=GOLD,
                            num_format=INT))
    ws.write_formula(21, 5, "=IFERROR(SUM($F$12:$F$21),0)",
                     k.f(font_name=EN, font_size=10, bold=True, font_color=WHITE,
                        bg_color=GOLD, align="center", border=1, border_color=GOLD,
                        num_format=PCT))
    for _c in (1, 2, 3, 4):
        ws.set_column(_c, _c, 15)

    ch = wb.add_chart({"type": "bar"})
    ch.add_series({
        "name": "مانده فعلی",
        "categories": ["Sum", 11, 0, 20, 0],
        "values": ["Sum", 11, 4, 20, 4],
        "fill": {"color": FOREST},
        "border": {"none": True},
        "data_labels": {"value": True, "num_format": "#,##0", "font": {"size": 7}},
    })
    ch.set_title({"name": "مانده هر بانک", "name_font": {"name": FA, "size": 11}})
    ch.set_legend({"none": True})
    ch.set_plotarea({"fill": {"color": "#FCFAF4"}})
    ch.set_chartarea({"border": {"color": LINE}, "fill": {"color": WHITE}})
    ch.set_size({"width": 470, "height": 230})
    ws.insert_chart("A25", ch, {"x_offset": 8, "y_offset": 6})
    return ws


# =====================================================================  Flow
def sheet_flow(wb, k):
    ws = wb.add_worksheet("Flow")
    last = 7
    k.setup(ws, last, 40, landscape=True, repeat=(10, 10), freeze=(12, 1))
    k.brand(ws, last, "جریان نقدی — ماه به ماه",
            "دریافت، پرداخت و خالص هر ماه شمسی  ·  ستون ماه از تاریخ تراکنش ساخته می‌شود", icon="chart")
    k.kpis(ws, 6, [
        ("بهترین ماه (خالص)", '=IFERROR(INDEX($A$13:$A$24,MATCH(MAX($E$13:$E$24),$E$13:$E$24,0)),"—")', "00", FOREST, "chart"),
        ("بیشترین دریافت", "=MAX($C$13:$C$24)", INT, GREEN, "transfer"),
        ("بیشترین پرداخت", "=MAX($D$13:$D$24)", INT, RED, "transfer"),
        ("میانگین ماهانه", "=IFERROR(AVERAGE($E$13:$E$24),0)", INT_RED, BLUE, "money"),
    ], last)
    heads = [("ماه", "Month", 9), ("تعداد", "Count", 9), ("دریافت", "Inflow", 16),
             ("پرداخت", "Outflow", 16), ("خالص", "Net", 16),
             ("تجمعی", "Cumulative", 16), ("نمودار", "Bar", 7)]
    for i, (fa, en, w) in enumerate(heads):
        ws.set_column(i, i, w)
        ws.write(10, i, fa, k.f(font_name=FA, font_size=9.5, bold=True, font_color=WHITE,
                                bg_color=FOREST, border=1, border_color=FOREST,
                                align="center"))
        ws.write(11, i, en, k.f(font_name=EN, font_size=6.8, font_color=SAGE_L,
                                bg_color=FOREST, align="center"))
    for m in range(12):
        r = 12 + m
        ws.write_number(r, 0, m + 1, k.f(font_name=EN, font_size=9, bold=True,
                                         font_color=FOREST, bg_color=SAND_2, border=1,
                                         border_color=LINE, align="center", num_format="00"))
        ws.write_formula(r, 1, f'=COUNTIFS(Post!$K$13:$K${12+PN},$A{r+1},'
                               f'Post!$N$13:$N${12+PN},1)',
                         k.f(font_name=EN, font_size=9, bg_color=WHITE, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(r, 2, f'=SUMIFS(Post!$I$13:$I${12+PN},Post!$K$13:$K${12+PN},$A{r+1},'
                               f'Post!$N$13:$N${12+PN},1)',
                         k.f(font_name=EN, font_size=9, bg_color=MINT, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(r, 3, f'=SUMIFS(Post!$J$13:$J${12+PN},Post!$K$13:$K${12+PN},$A{r+1},'
                               f'Post!$N$13:$N${12+PN},1)',
                         k.f(font_name=EN, font_size=9, bg_color=BLUSH, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(r, 4, f'=$C{r+1}-$D{r+1}',
                         k.f(font_name=EN, font_size=9, bold=True, bg_color=CREAM, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT_RED))
        ws.write_formula(r, 5, f'=IF($A{r+1}="","",SUM($E$13:$E{r+1}))',
                         k.f(font_name=EN, font_size=9, bg_color=WHITE, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT_RED))
        ws.write_formula(r, 6, f'=REPT("█",MIN(18,MAX(0,ROUND($C{r+1}/MAX($C$13:$C$24)*18,0))))',
                         k.f(font_name=EN, font_size=7, font_color=GREEN, bg_color=WHITE,
                             align="right"))
    ws.write(24, 0, "جمع سال",
             k.f(font_name=FA, font_size=10, bold=True, font_color=WHITE,
                 bg_color=GOLD, align="center", border=1, border_color=GOLD))
    for c in (1, 2, 3, 4):
        L = chr(65 + c)
        ws.write_formula(24, c, f"=SUM({L}13:{L}24)",
                         k.f(font_name=EN, font_size=10, bold=True, font_color=WHITE,
                            bg_color=GOLD, align="center", border=1, border_color=GOLD,
                            num_format=INT))

    ch = wb.add_chart({"type": "column"})
    ch.add_series({"name": "دریافت", "categories": ["Flow", 12, 0, 23, 0],
                   "values": ["Flow", 12, 2, 23, 2], "fill": {"color": SAGE},
                   "gap": 40})
    ch.add_series({"name": "پرداخت", "categories": ["Flow", 12, 0, 23, 0],
                   "values": ["Flow", 12, 3, 23, 3], "fill": {"color": TERRA}})
    ln = wb.add_chart({"type": "line"})
    ln.add_series({"name": "خالص", "categories": ["Flow", 12, 0, 23, 0],
                   "values": ["Flow", 12, 4, 23, 4],
                   "line": {"color": FOREST, "width": 2.5},
                   "marker": {"type": "circle", "size": 5, "fill": {"color": GOLD},
                              "border": {"color": WHITE}}})
    ch.combine(ln)
    ch.set_title({"name": "جریان نقدی ماهانه", "name_font": {"name": FA, "size": 11}})
    ch.set_legend({"position": "bottom", "font": {"name": FA, "size": 8}})
    ch.set_plotarea({"fill": {"color": "#FCFAF4"}})
    ch.set_size({"width": 620, "height": 280})
    ws.insert_chart("A27", ch, {"x_offset": 8, "y_offset": 6})
    return ws


# =====================================================================  Anal
def sheet_anal(wb, k):
    ws = wb.add_worksheet("Anal")
    last = 6
    k.setup(ws, last, 52, landscape=True, repeat=(7, 7), freeze=(10, 1))
    k.brand(ws, last, "تحلیل تراکنش‌ها",
            "بر پایه نوع عملیات، روش پرداخت و طرف حساب  ·  سه مدل تحلیل در یک صفحه", icon="chart")

    k.section(ws, 6, last, "۰۱", "تحلیل بر اساس نوع عملیات", "کدام عملیات بیشترین گردش را دارد؟")
    heads = [("نوع عملیات", "Operation", 26), ("تعداد", "Count", 9), ("دریافت", "In", 15),
             ("پرداخت", "Out", 15), ("خالص", "Net", 15), ("سهم پرداخت", "Share", 10),
             ("نمودار", "Bar", 8)]
    for i, (fa, en, w) in enumerate(heads):
        ws.set_column(i, i, w)
        ws.write(7, i, fa, k.f(font_name=FA, font_size=9.5, bold=True, font_color=WHITE,
                               bg_color=FOREST, border=1, border_color=FOREST, align="center"))
        ws.write(8, i, en, k.f(font_name=EN, font_size=6.8, font_color=SAGE_L,
                               bg_color=FOREST, align="center"))
    for i, (title, *_rest) in enumerate(BANK_OPS):
        r = 9 + i
        ws.write_string(r, 0, title, k.f(font_name=FA, font_size=9, bg_color=SAND_2,
                                         border=1, border_color=LINE, align="right"))
        ws.write_formula(r, 1, f'=COUNTIF(Txn!$D$13:$D${12+NR},$A{r+1})',
                         k.f(font_name=EN, font_size=9, bg_color=WHITE, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(r, 2, f'=SUMIFS(Txn!$M$13:$M${12+NR},Txn!$D$13:$D${12+NR},$A{r+1},'
                               f'Txn!$U$13:$U${12+NR},"IN")',
                         k.f(font_name=EN, font_size=9, bg_color=MINT, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(r, 3, f'=SUMIFS(Txn!$M$13:$M${12+NR},Txn!$D$13:$D${12+NR},$A{r+1},'
                               f'Txn!$U$13:$U${12+NR},"OUT")',
                         k.f(font_name=EN, font_size=9, bg_color=BLUSH, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(r, 4, f'=$C{r+1}-$D{r+1}',
                         k.f(font_name=EN, font_size=9, bold=True, bg_color=CREAM, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT_RED))
        ws.write_formula(r, 5, f'=IFERROR($D{r+1}/SUM($D$10:$D$25),0)',
                         k.f(font_name=EN, font_size=8.5, bg_color=WHITE, border=1,
                             border_color=LINE_SOFT, align="center", num_format=PCT))
        ws.write_formula(r, 6, f'=REPT("█",MIN(16,ROUND(MAX($C{r+1},$D{r+1})/'
                               f'MAX(1,MAX($C$10:$C$25),MAX($D$10:$D$25))*16,0)))',
                         k.f(font_name=EN, font_size=7, font_color=GOLD, bg_color=WHITE,
                             align="right"))

    ch = wb.add_chart({"type": "doughnut"})
    ch.add_series({
        "name": "پرداخت بر حسب عملیات",
        "categories": ["Anal", 9, 0, 24, 0],
        "values": ["Anal", 9, 3, 24, 3],
        "points": [{"fill": {"color": c}} for c in
                   (FOREST, SAGE, GOLD, TERRA, BLUE, AMBER, "#8E6C88", "#4F7C6C")],
        "data_labels": {"percentage": True, "font": {"size": 7}},
    })
    ch.set_title({"name": "سهم هر نوع عملیات از پرداخت‌ها", "name_font": {"name": FA, "size": 10}})
    ch.set_legend({"position": "right", "font": {"name": FA, "size": 7}})
    ch.set_size({"width": 380, "height": 250})
    ws.insert_chart("I8", ch, {"x_offset": 6, "y_offset": 4})

    # methods table
    k.section(ws, 27, last, "۰۲", "تحلیل بر اساس روش پرداخت", "نقد، چک، انتقال، کارت/پوز، حواله")
    mh = ["روش", "تعداد", "دریافت", "پرداخت", "خالص", "سهم"]
    for i, h in enumerate(mh):
        ws.write(28, i, h, k.f(font_name=FA, font_size=9, bold=True, font_color=WHITE,
                               bg_color=FOREST_2, border=1, border_color=FOREST_2,
                               align="center"))
    for i, m in enumerate(["نقد", "چک", "انتقال بانکی", "کارت/پوز", "حواله"]):
        r = 29 + i
        ws.write_string(r, 0, m, k.f(font_name=FA, font_size=9, bg_color=SAND_2, border=1,
                                     border_color=LINE, align="right"))
        ws.write_formula(r, 1, f'=COUNTIFS(Txn!$G$13:$G${12+NR},$A{r+1})',
                         k.f(font_name=EN, font_size=9, bg_color=WHITE, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(r, 2, f'=SUMIFS(Txn!$O$13:$O${12+NR},Txn!$G$13:$G${12+NR},$A{r+1})',
                         k.f(font_name=EN, font_size=9, bg_color=MINT, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(r, 3, f'=SUMIFS(Txn!$P$13:$P${12+NR},Txn!$G$13:$G${12+NR},$A{r+1})',
                         k.f(font_name=EN, font_size=9, bg_color=BLUSH, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(r, 4, f'=$C{r+1}-$D{r+1}',
                         k.f(font_name=EN, font_size=9, bold=True, bg_color=CREAM, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT_RED))
        ws.write_formula(r, 5, f'=IFERROR($D{r+1}/SUM($D$30:$D$34),0)',
                         k.f(font_name=EN, font_size=8.5, bg_color=WHITE, border=1,
                             border_color=LINE_SOFT, align="center", num_format=PCT))
    ch2 = wb.add_chart({"type": "pie"})
    ch2.add_series({"name": "روش پرداخت", "categories": ["Anal", 29, 0, 33, 0],
                    "values": ["Anal", 29, 3, 33, 3],
                    "points": [{"fill": {"color": c}} for c in (SAGE, GOLD, FOREST, TERRA, BLUE)],
                    "data_labels": {"percentage": True, "font": {"size": 8}}})
    ch2.set_title({"name": "ترکیب روش‌های پرداخت", "name_font": {"name": FA, "size": 10}})
    ch2.set_legend({"position": "right", "font": {"name": FA, "size": 7}})
    ch2.set_size({"width": 330, "height": 210})
    ws.insert_chart("I29", ch2, {"x_offset": 6, "y_offset": 4})

    # top counterparties
    k.section(ws, 36, last, "۰۳", "بزرگ‌ترین طرف حساب‌ها", "جمع گردش هر طرف حساب در دفتر تراکنش‌ها")
    th = ["کد حساب", "نام", "گروه", "جمع بدهکار", "جمع بستانکار", "خالص"]
    for i, h in enumerate(th):
        ws.write(37, i, h, k.f(font_name=FA, font_size=9, bold=True, font_color=WHITE,
                               bg_color=FOREST, border=1, border_color=FOREST, align="center"))
    for i in range(8):
        r = 38 + i
        for c in range(6):
            ws.write_blank(r, c, None, k.f(font_name=FA, font_size=9, bg_color=WHITE,
                                           border=1, border_color=LINE_SOFT))
        ws.write_formula(
            r, 0,
            f'=IFERROR(INDEX(Post!$D$13:$D${12+PN},MATCH(LARGE(Post!$I$13:$I${12+PN},1),'
            f'Post!$I$13:$I${12+PN},0)),"")',
            k.f(font_name=EN, font_size=8.5, bg_color=WHITE, border=1,
                border_color=LINE_SOFT, align="center"))
        ws.write_formula(r, 1, f'=IF($A{r+1}="","",IFERROR(VLOOKUP($A{r+1},AcctTbl,2,0),""))',
                         k.f(font_name=FA, font_size=9, bg_color=WHITE, border=1,
                             border_color=LINE_SOFT, align="right"))
        ws.write_formula(r, 2, f'=IF($A{r+1}="","",IFERROR(VLOOKUP($A{r+1},AcctTbl,3,0),""))',
                         k.f(font_name=FA, font_size=8.5, bg_color=WHITE, border=1,
                             border_color=LINE_SOFT, align="right"))
        ws.write_formula(r, 3, f'=IF($A{r+1}="","",SUMIF(Post!$D$13:$D${12+PN},$A{r+1},Post!$I$13:$I${12+PN}))',
                         k.f(font_name=EN, font_size=9, bg_color=WHITE, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(r, 4, f'=IF($A{r+1}="","",SUMIF(Post!$D$13:$D${12+PN},$A{r+1},Post!$J$13:$J${12+PN}))',
                         k.f(font_name=EN, font_size=9, bg_color=WHITE, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(r, 5, f'=IF($A{r+1}="","",$D{r+1}-$E{r+1})',
                         k.f(font_name=EN, font_size=9, bold=True, bg_color=CREAM, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT_RED))
    k.note(ws, 47, last,
           "فهرست بالا با LARGE مرتب می‌شود؛ برای دیدن ریز هر طرف حساب، شیت Post را فیلتر کنید یا از فایل گزارش (03) استفاده کنید.",
           bg=SAND_2, icon="search")
    return ws


# =====================================================================  Card
def sheet_card(wb, k, banks):
    ws = wb.add_worksheet("Card")
    last = 6
    k.setup(ws, last, 34, landscape=True, repeat=(10, 10), freeze=(12, 2))
    k.brand(ws, last, "کارت به کارت و پز — تحلیل بانکی",
            "هر واریزی کارت/پوز کدام بانک را شارژ کرده است؟  ·  پایه کنترل با باسکول", icon="card")
    k.kpis(ws, 6, [
        ("کل کارت/پوز", '=SUM($C$13:$C$22)', INT, FOREST, "card"),
        ("کل پز باسکول", '=SUM($D$13:$D$22)', INT, GOLD, "pos"),
        ("کارت به کارت سایر", "=$A$8-$C$8", INT, TERRA, "transfer"),
        ("سهم ۲۰۰۶", "=IFERROR($C$8/$A$8,0)", PCT, GREEN, "money"),
    ], last)
    heads = [("بانک", "Bank", 26), ("تعداد", "Count", 9), ("کارت/پوز", "Card", 16),
             ("از آن: پز باسکول", "POS", 16), ("سایر کارت‌ها", "Other", 16),
             ("سهم", "Share", 9), ("هشدار", "Alert", 12)]
    for i, (fa, en, w) in enumerate(heads):
        ws.set_column(i, i, w)
        ws.write(10, i, fa, k.f(font_name=FA, font_size=9.5, bold=True, font_color=WHITE,
                                bg_color=FOREST, border=1, border_color=FOREST, align="center"))
        ws.write(11, i, en, k.f(font_name=EN, font_size=6.8, font_color=SAGE_L,
                                bg_color=FOREST, align="center"))
    for i, (code, title, opening, accode, active, note) in enumerate(banks):
        r = 12 + i
        ws.write_formula(r, 0, f'=INDEX(BankTbl,{i+1},2)',
                         k.f(font_name=FA, font_size=9.5, bold=True, font_color=FOREST,
                             bg_color=SAND_2, border=1, border_color=LINE, align="right"))
        ws.write_formula(
            r, 1,
            f'=COUNTIFS(Post!$D$13:$D${12+PN},INDEX(BankTbl,{i+1},4),Post!$G$13:$G${12+PN},"*پز*")'
            f'+COUNTIFS(Post!$D$13:$D${12+PN},INDEX(BankTbl,{i+1},4),Post!$G$13:$G${12+PN},"*دریافت*")',
            k.f(font_name=EN, font_size=9, bg_color=WHITE, border=1,
                border_color=LINE_SOFT, align="center"))
        ws.write_formula(
            r, 2,
            f'=SUMIFS(Txn!$O$13:$O${12+NR},Txn!$E$13:$E${12+NR},INDEX(BankTbl,{i+1},1),'
            f'Txn!$G$13:$G${12+NR},"کارت/پوز")',
            k.f(font_name=EN, font_size=9, bg_color=MINT, border=1,
                border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(
            r, 3,
            f'=SUMIFS(Txn!$O$13:$O${12+NR},Txn!$E$13:$E${12+NR},INDEX(BankTbl,{i+1},1),'
            f'Txn!$D$13:$D${12+NR},"واریز پز باسکول (۲۰۰۶)")',
            k.f(font_name=EN, font_size=9, bg_color=GOLD_L, border=1,
                border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(r, 4, f'=$C{r+1}-$D{r+1}',
                         k.f(font_name=EN, font_size=9, bold=True, bg_color=CREAM, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT_RED))
        ws.write_formula(r, 5, f'=IFERROR($C{r+1}/SUM($C$13:$C$22),0)',
                         k.f(font_name=EN, font_size=9, bg_color=WHITE, border=1,
                             border_color=LINE_SOFT, align="center", num_format=PCT))
        ws.write_formula(r, 6, f'=IF($C{r+1}=0,"—",IF($D{r+1}=0,"ℹ بدون پز","✓"))',
                         k.f(font_name=FA, font_size=9, bg_color=WHITE, border=1,
                             border_color=LINE_SOFT, align="center"))
    ch = wb.add_chart({"type": "column"})
    ch.add_series({"name": "کارت/پوز", "categories": ["Card", 12, 0, 21, 0],
                   "values": ["Card", 12, 2, 21, 2], "fill": {"color": SAGE}, "gap": 60})
    ch.add_series({"name": "پز باسکول", "categories": ["Card", 12, 0, 21, 0],
                   "values": ["Card", 12, 3, 21, 3], "fill": {"color": GOLD}})
    ch.set_title({"name": "کارت/پوز هر بانک", "name_font": {"name": FA, "size": 11}})
    ch.set_legend({"position": "bottom", "font": {"name": FA, "size": 8}})
    ch.set_plotarea({"fill": {"color": "#FCFAF4"}})
    ch.set_size({"width": 560, "height": 250})
    ws.insert_chart("A24", ch, {"x_offset": 8, "y_offset": 6})
    k.note(ws, 34, last,
           "پز باسکول همان واریزی سرجمع روز بعد است. اگر جمع پز این صفحه با جمع پز شیت POS در 01_Entry یکی نبود، "
           "یعنی یک واریزی ثبت نشده یا یک تراکنش به بانک اشتباه زده شده است.", bg=SAND, icon="warn")
    return ws


# =====================================================================  Chq
def sheet_chq(wb, k):
    ws = wb.add_worksheet("Chq")
    last = 11
    k.setup(ws, last, 12 + 60 + 3, landscape=True, repeat=(11, 11), freeze=(12, 2))
    k.brand(ws, last, "دفتر چک‌ها",
            "چک‌های دریافتی و پرداختی  ·  سررسید، وضعیت و مبلغ در یک نگاه", icon="cheque")
    k.kpis(ws, 6, [
        ("تعداد چک", "=COUNTA($C$13:$C$72)", INT, FOREST, "cheque"),
        ("در جریان", '=SUMIFS($I$13:$I$72,$J$13:$J$72,"در جریان")', INT, GOLD, "calendar"),
        ("پاس شده", '=SUMIFS($I$13:$I$72,$J$13:$J$72,"پاس شده")', INT, GREEN, "check"),
        ("برگشتی", '=SUMIFS($I$13:$I$72,$J$13:$J$72,"برگشتی")', INT, RED, "warn"),
        ("جمع کل", "=SUM($I$13:$I$72)", INT, BLUE, "money"),
    ], last)
    heads = [("ردیف", "No.", 6), ("تاریخ صدور", "Issued", 12), ("سررسید", "Due", 12),
             ("شماره چک", "No.", 12), ("نوع", "Type", 11), ("کد طرف", "Party", 11),
             ("نام طرف", "Name", 24), ("بانک", "Bank", 18), ("مبلغ", "Amount", 15),
             ("وضعیت", "Status", 13), ("تاریخ پاس", "Cleared", 12), ("یادداشت", "Note", 22)]
    k.thead(ws, 11, heads, FOREST)
    f_in = k.f(font_name=EN, font_size=9, bg_color="#FFF4D6", border=1,
               border_color=GOLD_L, align="center", valign="vcenter")
    f_in_t = k.f(font_name=FA, font_size=9, bg_color="#FFF4D6", border=1,
                 border_color=GOLD_L, align="right", valign="vcenter")
    f_auto = k.f(font_name=FA, font_size=9, bg_color="#F6F5F0", border=1,
                 border_color=LINE_SOFT, align="right", valign="vcenter")
    for i in range(60):
        r = 12 + i
        x = r + 1
        ws.set_row(r, 21)
        ws.write_formula(r, 0, f'=IF($C{x}="","",COUNTA($C$13:$C{x}))',
                         k.f(font_name=EN, font_size=8.5, bg_color="#F6F5F0", border=1,
                             border_color=LINE_SOFT, align="center"))
        for c in (1, 2, 3):
            ws.write_blank(r, c, None, f_in)
        ws.write_blank(r, 4, None, f_in)
        ws.write_blank(r, 5, None, f_in)
        ws.write_formula(r, 6, f'=IF($F{x}="","",IFERROR(VLOOKUP($F{x},AcctTbl,2,0),""))', f_auto)
        ws.write_blank(r, 7, None, f_in_t)
        ws.write_blank(r, 8, None, k.f(font_name=EN, font_size=9, bg_color="#FFF4D6", border=1,
                                       border_color=GOLD_L, align="center", num_format=INT))
        ws.write_blank(r, 9, None, f_in)
        ws.write_blank(r, 10, None, f_in)
        ws.write_blank(r, 11, None, f_in_t)
    for i, row in enumerate(CHEQUES):
        r = 12 + i
        issued, due, num, typ, party, bank, amount, status, cleared, note = row
        ws.write_string(r, 1, issued, f_in)
        ws.write_string(r, 2, due, f_in)
        ws.write_string(r, 3, num, f_in)
        ws.write_string(r, 4, typ, f_in)
        ws.write_string(r, 5, party, f_in)
        ws.write_string(r, 7, bank, f_in_t)
        ws.write_number(r, 8, amount, k.f(font_name=EN, font_size=9, bg_color="#FFF4D6",
                                          border=1, border_color=GOLD_L, align="center",
                                          valign="vcenter", num_format=INT))
        ws.write_string(r, 9, status, f_in)
        if cleared:
            ws.write_string(r, 10, cleared, f_in)
        if note:
            ws.write_string(r, 11, note, f_in_t)
    ws.data_validation(12, 4, 71, 4, {"validate": "list", "source": "=ListChequeTypes",
                                     "error_type": "stop"})
    ws.data_validation(12, 7, 71, 7, {"validate": "list", "source": "=ListBanks",
                                      "error_type": "warning"})
    ws.data_validation(12, 9, 71, 9, {"validate": "list", "source": "=ListChequeStatus"})
    ws.data_validation(12, 5, 71, 5, {"validate": "list", "source": "=ListAllCodes",
                                      "error_type": "warning"})
    ws.conditional_format(12, 9, 71, 9, {
        "type": "text", "criteria": "containing", "value": "پاس",
        "format": k.f(bg_color=MINT, font_color=GREEN, bold=True)})
    ws.conditional_format(12, 9, 71, 9, {
        "type": "text", "criteria": "containing", "value": "برگشتی",
        "format": k.f(bg_color="#FBE3D8", font_color=TERRA, bold=True)})
    ws.conditional_format(12, 9, 71, 9, {
        "type": "text", "criteria": "containing", "value": "در جریان",
        "format": k.f(bg_color=SAND, font_color=GOLD, bold=True)})
    ws.autofilter(11, 0, 71, 11)
    return ws


# =====================================================================  Home
def sheet_home(wb, k, acc, banks):
    ws = wb.add_worksheet("Home")
    last = 8
    k.setup(ws, last, 40, landscape=True, repeat=None, freeze=None)
    k.brand(ws, last, "فایل تراکنش‌های بانکی",
            "ثبت، کنترل و تحلیل گردش حساب‌های شرکت  ·  همه کشوها از فایل مادر تعریف شده‌اند", icon="bank")
    k.kpis(ws, 6, [
        ("تراکنش‌ها", "=COUNTA(Txn!$B$13:$B$152)", INT, FOREST, "list"),
        ("دریافت", "=SUM(Txn!$O$13:$O$152)", INT, GREEN, "transfer"),
        ("پرداخت", "=SUM(Txn!$P$13:$P$152)", INT, RED, "transfer"),
        ("مانده بانک‌ها", "=SUM(Sum!$E$12:$E$21)", INT, BLUE, "bank"),
    ], last)
    k.section(ws, 8, last, "۰۱", "شیت‌های این فایل", "روی هر کارت کلیک کنید")
    k.nav(ws, 9, [
        ("تراکنش‌ها  Txn", "ثبت تراکنش با کشوی بانک، کد طرف حساب و انتقال داخلی", "internal:Txn!A1", "bank", FOREST),
        ("مانده بانک‌ها  Sum", "مانده هر حساب + نمودار مقایسه بانک‌ها", "internal:Sum!A1", "money", GOLD),
        ("جریان نقدی  Flow", "دریافت/پرداخت ماهانه با نمودار ترکیبی", "internal:Flow!A1", "chart", SAGE),
        ("تحلیل  Anal", "بر پایه نوع عملیات، روش پرداخت و طرف حساب", "internal:Anal!A1", "gauge", BLUE),
        ("کارت/پوز  Card", "کارت‌به‌کارت و پز باسکول به تفکیک بانک", "internal:Card!A1", "card", TERRA),
        ("چک‌ها  Chq", "سررسید، وضعیت و مبلغ چک‌های دریافتی/پرداختی", "internal:Chq!A1", "cheque", "#8E6C88"),
    ], last, per_row=3)

    row = 21
    k.section(ws, row, last, "۰۲", "گردش کار روزانه", "پنج گام ساده")
    steps = [
        ("۱", "صورت‌حساب بانک را باز کنید؛ هر ردیف را در شیت Txn ثبت کنید."),
        ("۲", "نوع عملیات را انتخاب کنید؛ جهت بدهکار/بستانکار خودکار تعیین می‌شود."),
        ("۳", "بانک عامل را از کشو بزنید و کد طرف حساب را تایپ کنید (نام خودکار می‌آید)."),
        ("۴", "اگر پول بین دو حساب خودمان جابه‌جا شده، «انتقال داخلی» را بزنید و بانک مقصد را هم انتخاب کنید."),
        ("۵", "آخر روز شیت Sum و Card را ببینید؛ اگر همه ردیف‌ها ✓ بودند، فایل‌های گزارش را باز کنید."),
    ]
    r = row + 1
    for n, txt in steps:
        ws.write(r, 0, n, k.f(font_name=EN, font_size=10, bold=True, font_color=WHITE,
                              bg_color=GOLD, align="center", valign="vcenter"))
        ws.merge_range(r, 1, r, last, txt,
                       k.f(font_name=FA, font_size=10, color=INK) if False else
                       k.f(font_name=FA, font_size=10, font_color=INK, align="right",
                           valign="vcenter", border=1, border_color=LINE_SOFT))
        ws.set_row(r, 24)
        r += 1
    k.note(ws, r + 1, last,
           "قاعده طلایی: هر بانکی که در فهرست بانک‌های شرکت نیست، اجازه ثبت ندارد. اگر حساب جدیدی باز کردید، "
           "اول آن را در فایل مادر (شیت Bnk) اضافه کنید و بعد این فایل را باز کنید.", bg=MINT, icon="flag")
    return ws


# =====================================================================  Help
def sheet_help(wb, k, acc, banks):
    ws = wb.add_worksheet("Help")
    last = 8
    k.setup(ws, last, 96, landscape=False, repeat=(7, 7), freeze=None)
    k.brand(ws, last, "راهنمای کامل فایل تراکنش‌ها",
            "از تعریف کشوها تا قواعد بدهکار/بستانکار و روش‌های کنترل", icon="help")

    blocks = [
        ("۰۱", "این فایل چه می‌کند؟", [
            ("p", "فایل 02_Bank محل ثبت تراکنش‌های بانکی است. برخلاف بقیه فایل‌ها که داده را از 01_Entry می‌خوانند، "
                  "این فایل منبع مستقل ورود اطلاعات بانکی است، چون صورت‌حساب بانک از بیرون می‌آید."),
            ("b", "خروجی: شیت Post (دفتر تراکنش‌ها) که همه تحلیل‌ها و مانده‌ها از آن ساخته می‌شود."),
            ("b", "هر تراکنش دو خط دفتر می‌سازد: خط بانک و خط طرف حساب."),
            ("b", "کد حساب‌های بانکی در همه فایل‌ها BNxx است (BN01…BN10). در دفتر روزنامه (01_Entry) هیچ سندی "
                  "مستقیم روی BNxx زده نمی‌شود: هر حرکت بانکی از حساب واسط X011 عبور می‌کند و مانده واقعی "
                  "هر بانک فقط از همین فایل (شیت Post) خوانده می‌شود."),
            ("b", "دفتر روزنامه تنها دفتر مالی است: هر رویداد اقتصادی (فروش، هزینه، بدهی) باید در 01_Entry ثبت شود و "
                  "این فایل حرکت واقعی پول را. برای هزینه‌ای که از بانک پرداخت می‌شود، در 01_Entry یک «پرداخت» با "
                  "روش «انتقال بانکی» و حساب هزینه بزنید و همین‌جا همان مبلغ را با نوع «تسویه پرداخت (واسط)» ثبت کنید."),
        ]),
        ("۰۲", "کشوها از کجا می‌آیند؟", [
            ("p", "همه فهرست‌های کشویی در فایل مادر (00_Master) تعریف شده‌اند و در شیت Sync این فایل کپی شده‌اند:"),
            ("b", "بانک عامل و بانک مقصد ← شیت Bnk فایل مادر."),
            ("b", "نوع عملیات بانکی ← جدول «Bank Operations» در شیت List فایل مادر (۱۸ نوع، قابل افزودن)."),
            ("b", "روش پرداخت و وضعیت چک ← شیت List."),
            ("b", "شخص باسکول دیگران ← حساب‌های BD001…BD020 در شیت Acc."),
            ("b", "کد طرف حساب یک کشوی «هشدار» است: می‌توانید کد جدید را تایپ کنید، ولی اگر کد در فایل مادر نباشد ⚠ می‌گیرید."),
        ]),
        ("۰۳", "انتقال داخلی بین بانک‌ها", [
            ("p", "وقتی پول از یک حساب شرکت به حساب دیگر منتقل می‌شود، این یک تراکنش واقعی نیست و تراز شرکت تغییر نمی‌کند."),
            ("b", "نوع عملیات = «انتقال داخلی بین بانک‌ها»."),
            ("b", "بانک عامل = بانک مبدأ (پول از آن خارج می‌شود)."),
            ("b", "بانک مقصد = بانکی که پول به آن واریز می‌شود."),
            ("b", "نتیجه: خط اول بستانکار بانک مبدأ، خط دوم بدهکار بانک مقصد؛ جمع شرکت صفر می‌ماند."),
            ("b", "دفتر روزنامه برای «انتقال داخلی» هیچ سندی نمی‌سازد، چون پولی از شرکت خارج نمی‌شود؛ فقط همین فایل "
                  "هر دو طرف را ثبت می‌کند. در شیت پرداخت 01_Entry هم روش «انتقال داخلی» را انتخاب کنید تا سند اضافه زده نشود."),
        ]),
        ("۰۴", "انعام مصطفی", [
            ("p", "حساب انعام مصطفی کد X001 دارد و در فایل مادر تعریف شده است. پرداخت به او با نوع «پرداخت انعام مصطفی» ثبت می‌شود: "
                  "بانک عامل بستانکار و حساب X001 بدهکار می‌شود؛ یعنی بدهی شرکت به ایشان تسویه می‌شود."),
            ("b", "برای دیدن مانده انعام مصطفی، در شیت Post ستون کد حساب را روی X001 فیلتر کنید."),
            ("b", "کنترل خودکار ۲۰۰۶ ↔ مصطفی در شیت POS فایل 01_Entry (کنترل مانده واسط X011) انجام می‌شود و باید صفر بماند."),
        ]),
        ("۰۵", "باسکول دیگران", [
            ("p", "دریافت هزینه باسکول دیگران: بانک عامل بدهکار (پول وارد شده) و حساب همان شخص بستانکار؛ یعنی شرکت به او بدهکار می‌شود."),
            ("b", "شخص را از کشوی «شخص باسکول دیگران» انتخاب کنید؛ کد (BDxxx) در ستون مخفی محاسبه می‌شود."),
            ("b", "پرداخت به باسکول دیگران برعکس عمل می‌کند و بدهی را کم می‌کند."),
            ("b", "گزارش بدهی هر شخص: شیت Anal یا فیلتر حساب BDxxx در شیت Post."),
        ]),
        ("۰۶", "چک‌ها", [
            ("p", "شیت Chq دفتر چک است: شماره، سررسید، مبلغ، وضعیت (در جریان/پاس شده/برگشتی/باطل)."),
            ("b", "وقتی چک پاس شد، در Txn یک ردیف «پاس چک» ثبت کنید تا حساب X010 و بانک به‌روز شود."),
            ("b", "چک برگشتی با نوع «برگشت چک دریافتی» ثبت می‌شود."),
        ]),
        ("۰۷", "کنترل‌ها و هشدارها", [
            ("b", "ستون وضعیت Txn: ⚠ مبلغ / ⚠ بانک عامل / ⚠ بانک مقصد / ⚠ طرف حساب."),
            ("b", "در انتهای Txn جمع بدهکار و بستانکار باید برابر باشند؛ کادر «تراز» سبز می‌شود."),
            ("b", "شیت Card: واریزی پز باسکول هر بانک با پز ثبت‌شده در دفتر روزنامه مقایسه می‌شود؛ این دو باید برابر باشند."),
            ("b", "اگر نوع عملیات را از گروه «تسویه…» انتخاب نکنید (مثلاً «پرداخت هزینه» با کد طرف حساب)، آن مبلغ "
                  "در دفتر روزنامه سند ندارد و جمع واسط X011 صفر نمی‌شود؛ شیت 05_Bal (شیت Audit) این اختلاف را نشان می‌دهد."),
        ]),
        ("۰۸", "چاپ و صفحه‌آرایی", [
            ("p", "همه شیت‌ها راست‌به‌چپ و آماده چاپ A4 هستند: تکرار سرستون، جا شدن در عرض یک صفحه، سربرگ و فوتر با شماره صفحه."),
            ("b", "برای چاپ فقط ناحیه داده، از سربرگ Page Layout ← Print Area استفاده کنید (از پیش تنظیم شده)."),
        ]),
        ("۰۹", "ارتباط با بقیه فایل‌ها", [
            ("b", "00_Master ← تعریف بانک‌ها، حساب‌ها و انواع عملیات."),
            ("b", "01_Entry ← عملیات باسکول و پرداخت‌های نقدی/چکی؛ پز باسکول آنجا ثبت می‌شود."),
            ("b", "08_Dash ← داشبوردی که هم از این فایل و هم از 01_Entry می‌خواند."),
            ("b", "05_Bal ← مانده هر بانک در شیت IBank از همین فایل (Post) خوانده می‌شود؛ جمع آن با شیت SUM همین فایل و داشبورد یکی است."),
            ("b", "03_Report / 04_Inv ← گزارش‌ها و صورت‌حساب‌ها."),
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
                ws.set_row(r, 28)
                r += 1
        r += 1
    ws.print_area(0, 0, r, last)
    return ws


def build(master_path, out_path, demo=None):
    acc, banks, gl, bd = master_data(master_path)
    global DEMO
    if demo is not None:
        DEMO = list(demo)
    wb = xlsxwriter.Workbook(out_path, {"default_format_properties": {
        "font_name": FA, "font_size": 10}})
    wb.set_properties({
        "title": "مازندشت — تراکنش‌های بانکی",
        "subject": "ثبت و تحلیل تراکنش بانکی",
        "author": "MAZANDASHT",
        "company": "سورتینگ و بسته‌بندی مرکبات مازندشت",
        "comments": "نسخه ۰۳ — Excel 2013",
    })
    k = Kit(wb)
    sheet_home(wb, k, acc, banks)
    sheet_txn(wb, k)
    sheet_post(wb, k)
    sheet_sum(wb, k, banks)
    sheet_flow(wb, k)
    sheet_anal(wb, k)
    sheet_card(wb, k, banks)
    sheet_chq(wb, k)
    sheet_sync(wb, k, acc, banks, gl, bd)
    sheet_help(wb, k, acc, banks)
    wb.worksheets()[0].activate()
    wb.close()
    return out_path
