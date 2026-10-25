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
NR = 5000         # transaction rows  (12 .. 12+NR-1)
PN = NR * 2       # journal lines     (2 per transaction)
NCHQ = 300        # cheque rows

DEMO = [
    ("1405/07/01", "TB-001", "واریز پز باسکول (۲۰۰۶)", "BN02", "", "کارت/پوز", "",
     "واریز سرجمع پز ۱۴۰۵/۰۶/۳۱", "", 120000000, "", "فروش مرکبات"),
    ("1405/07/01", "TB-002", "دریافت از مشتری", "BN01", "", "انتقال بانکی", "C001",
     "علی‌الحساب خرید مرکبات", "TR-8841", 150000000, "", "فروش مرکبات"),
    ("1405/07/02", "TB-003", "پرداخت انعام مصطفی", "BN02", "", "کارت/پوز", "",
     "انعام هفته اول", "", 5000000, "", "انعام مصطفی"),
    ("1405/07/02", "TB-004", "انتقال داخلی بین بانک‌ها", "BN01", "BN03", "انتقال بانکی", "",
     "تقویت حساب ملت", "TR-8842", 200000000, "", "انتقال بین حساب‌های شرکت"),
    ("1405/07/03", "TB-005", "پرداخت به باغدار", "BN03", "", "حواله", "C002",
     "تسویه بخشی از حساب باغدار", "TR-8843", 60000000, "", "پرداخت به باغدار"),
    ("1405/07/03", "TB-006", "دریافت هزینه باسکول دیگران", "BN02", "", "کارت/پوز", "",
     "هزینه باسکول میدان بار", "", 1500000, "BD001 – باسکول میدان بار (نمونه)",
     "هزینه باسکول دیگران"),
    ("1405/07/04", "TB-007", "واریز پز باسکول (۲۰۰۶)", "BN02", "", "کارت/پوز", "",
     "واریز پز ۱۴۰۵/۰۷/۰۳", "", 95000000, "", "فروش مرکبات"),
    ("1405/07/05", "TB-008", "پرداخت هزینه", "BN01", "", "انتقال بانکی", "X012",
     "هزینه حمل و تخلیه", "", 12000000, "", "حمل و تخلیه"),
    ("1405/07/06", "TB-009", "پاس چک ما (چک پرداختی شرکت)", "BN01", "", "پاس چک", "",
     "چک ۴۵۲۱ بابت تعمیر دستگاه سورتینگ", "4521", 40000000, "",
     "تعمیر و نگهداری دستگاه سورتینگ"),
    ("1405/07/07", "TB-010", "واریز بارفروشان", "BN03", "", "انتقال بانکی", "C003",
     "واریز بابت فاکتور بار", "TR-8844", 88000000, "", "واریز بارفروشان"),
    ("1405/07/08", "TB-011", "واریز نقدی به حساب", "BN01", "", "نقد", "",
     "واریز نقدی صندوق", "", 25000000, "", "سایر درآمد"),
    ("1405/07/09", "TB-012", "پرداخت به مشتری", "BN02", "", "کارت/پوز", "S001",
     "پرداخت به فروشنده سبد", "", 33000000, "", "خرید سبد"),
    ("1405/07/10", "TB-013", "واریز پز باسکول (۲۰۰۶)", "BN02", "", "کارت/پوز", "",
     "واریز پز ۱۴۰۵/۰۷/۰۹", "", 76000000, "", "فروش مرکبات"),
    ("1405/07/11", "TB-014", "پرداخت به باسکول دیگران", "BN02", "", "کارت/پوز", "",
     "تسویه با باسکول میدان بار", "", 1000000, "BD001 – باسکول میدان بار (نمونه)",
     "تسویه باسکول دیگران"),
    ("1405/07/12", "TB-015", "سایر دریافت", "BN03", "", "انتقال بانکی", "C005",
     "دریافت متفرقه", "", 9500000, "", "سایر درآمد"),
    ("1405/07/08", "TB-016", "پرداخت به حساب شخصی رئیس", "BN01", "", "کارت/پوز", "",
     "خرید لوازم خانگی شخصی", "", 18000000, "", "خرید وسایل شخصی"),
    ("1405/07/10", "TB-017", "پرداخت به حساب شخصی رئیس", "BN03", "", "انتقال بانکی", "",
     "سرویس و لاستیک خودرو شخصی", "", 9500000, "", "سوخت و تعمیر خودرو شخصی"),
    ("1405/07/12", "TB-018", "دریافت از حساب شخصی رئیس", "BN02", "", "کارت/پوز", "",
     "برگشت پول شخصی رئیس به شرکت", "", 12000000, "", "برگشت وجه شخصی به شرکت"),
    ("1405/07/13", "TB-019", "کارمزد و هزینه بانکی", "BN01", "", "انتقال بانکی", "",
     "کارمزد ماهانه حساب", "", 450000, "", "کارمزد بانکی"),
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
    for r in range(10, 300):
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


def read_classes(master_path):
    """Classification table (00_Master → Cls); falls back to the built-in list."""
    try:
        wb = openpyxl.load_workbook(master_path, read_only=True)
    except Exception:
        return list(CLASSES)
    if "Cls" not in wb.sheetnames:
        wb.close()
        return list(CLASSES)
    ws = wb["Cls"]
    out = []
    for r in range(10, 400):
        title = ws.cell(r, 2).value
        if not title:
            continue
        out.append((str(ws.cell(r, 1).value or f"TC{r-9:02d}"), str(title),
                    str(ws.cell(r, 3).value or "شرکتی"), str(ws.cell(r, 4).value or "سایر حساب")))
    wb.close()
    return out or list(CLASSES)


# =====================================================================  Sync
def sheet_sync(wb, k, acc, banks, gl, bd, classes=()):
    ws = wb.add_worksheet("Sync")
    k.setup(ws, 28, max(262, 12 + len(classes or CLASSES)), landscape=True,
              repeat=(8, 9), freeze=(9, 3))
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

    ws.merge_range(0, 0, 0, 28, "داده‌های ثابت — کپی قفل‌شده از فایل مادر 00_Master.xlsx",
                   k.f(font_name=FA_TITLE, font_size=15, bold=True, font_color=FOREST,
                       bg_color=SAND_2, align="right", valign="vcenter"))
    ws.merge_range(1, 0, 1, 28,
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
    for i, (title, direction, mode, fixed, note) in enumerate(TXN_OPS):
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

    # ---- classifications Y..AC  (title, code, scope, group, note)
    for c_, w in ((24, 34), (25, 8), (26, 13), (27, 22), (28, 40)):
        ws.set_column(c_, c_, w)
    head(8, 24, 28, "طبقه‌بندی تراکنش — ماهیت و دامنه", "TRANSACTION CLASSES", TERRA)
    for i, (code, title, scope, group) in enumerate(classes or CLASSES):
        r = 9 + i
        ws.write(r, 24, title, f_t)
        ws.write(r, 25, code, f_c)
        ws.write(r, 26, scope, f_c)
        ws.write(r, 27, group, f_t)
        ws.write(r, 28, "", f_t)
    head(8, 22, 22, "گروه‌های حساب", "GROUPS", GOLD)
    for i, g in enumerate(gl[:24]):
        ws.write(9 + i, 22, g, f_t)

    # ---- names
    wb.define_name("BankTbl", "=Sync!$A$10:$F$19")
    wb.define_name("BankAccts", "=Sync!$D$10:$D$19")
    wb.define_name("ListBanks", "=Sync!$A$10:$A$19")
    n_acc = 9 + len(acc)
    wb.define_name("AcctTbl", f"=Sync!$H$10:$L${n_acc}")
    wb.define_name("ListAllCodes", f"=Sync!$H$10:$H${n_acc}")
    wb.define_name("ListAcctNames", f"=Sync!$I$10:$I${n_acc}")
    n_ops = 9 + len(TXN_OPS)
    n_cls = 9 + len(classes or CLASSES)
    wb.define_name("BankOpTbl", f"=Sync!$N$10:$S${n_ops}")
    wb.define_name("ListTxnOps", f"=Sync!$N$10:$N${n_ops}")
    wb.define_name("ListBankOps", f"=Sync!$N$10:$N${n_ops}")
    wb.define_name("ClassTbl", f"=Sync!$Y$10:$AC${n_cls}")
    wb.define_name("ListClasses", f"=Sync!$Y$10:$Y${n_cls}")
    wb.define_name("ListPayMethods", "=Sync!$T$10:$T$14")
    wb.define_name("ListChequeStatus", "=Sync!$U$10:$U$13")
    wb.define_name("ListOtherScales", "=Sync!$V$10:$V$29")
    wb.define_name("ListGroups", "=Sync!$W$10:$W$33")
    wb.define_name("ListChequeTypes", "=Sync!$X$10:$X$11")

    ws.protect("", {"select_locked_cells": True, "select_unlocked_cells": False,
                    "format_cells": False, "format_columns": False, "format_rows": False})
    return ws


# =====================================================================  Txn

CLASSES = [
    # (code, title, scope, group)
    ("TC01", "خرید مرکبات", "شرکتی", "باغدار"),
    ("TC02", "فروش مرکبات", "شرکتی", "بارفروش"),
    ("TC03", "واریز بارفروشان", "شرکتی", "بارفروش"),
    ("TC04", "دریافتی از مشتری", "شرکتی", "بارفروش"),
    ("TC05", "پرداخت به باغدار", "شرکتی", "باغدار"),
    ("TC06", "خرید سبد", "شرکتی", "فروشنده سبد"),
    ("TC07", "فروش سبد", "شرکتی", "فروشنده سبد"),
    ("TC08", "حمل و تخلیه", "شرکتی", "سایر حساب"),
    ("TC09", "کارمزد بانکی", "شرکتی", "سایر حساب"),
    ("TC10", "قبض آب، برق و گاز", "شرکتی", "سایر حساب"),
    ("TC11", "حقوق و دستمزد", "شرکتی", "کارمند"),
    ("TC12", "بیمه تأمین اجتماعی", "شرکتی", "کارمند"),
    ("TC13", "تعمیر و نگهداری دستگاه سورتینگ", "شرکتی", "سایر حساب"),
    ("TC14", "تعمیر و نگهداری خط بسته‌بندی", "شرکتی", "سایر حساب"),
    ("TC15", "تعمیر و نگهداری سردخانه", "شرکتی", "سایر حساب"),
    ("TC16", "قطعات و لوازم یدکی", "شرکتی", "سایر حساب"),
    ("TC17", "سوخت و انرژی", "شرکتی", "سایر حساب"),
    ("TC18", "اجاره انبار و زمین", "شرکتی", "سایر حساب"),
    ("TC19", "انعام مصطفی", "شرکتی", "سایر حساب"),
    ("TC20", "هزینه باسکول دیگران", "شرکتی", "باسکول دیگران"),
    ("TC21", "تسویه باسکول دیگران", "شرکتی", "باسکول دیگران"),
    ("TC22", "اقساط و تسهیلات", "شرکتی", "وام و تسهیلات"),
    ("TC23", "دریافت تسهیلات", "شرکتی", "وام و تسهیلات"),
    ("TC24", "مالیات و عوارض", "شرکتی", "سایر حساب"),
    ("TC25", "سرمایه و آورده صاحبان", "شرکتی", "سرمایه و آورده صاحبان"),
    ("TC26", "انتقال بین حساب‌های شرکت", "شرکتی", "سایر حساب"),
    ("TC27", "سایر درآمد", "شرکتی", "سایر حساب"),
    ("TC28", "سایر هزینه", "شرکتی", "سایر حساب"),
    ("TC29", "هزینه شخصی رئیس", "شخصی رئیس", "سایر حساب"),
    ("TC30", "خرید وسایل شخصی", "شخصی رئیس", "سایر حساب"),
    ("TC31", "فروش وسایل شخصی", "شخصی رئیس", "سایر حساب"),
    ("TC32", "خرید خودرو", "شخصی رئیس", "سایر حساب"),
    ("TC33", "فروش خودرو", "شخصی رئیس", "سایر حساب"),
    ("TC34", "سوخت و تعمیر خودرو شخصی", "شخصی رئیس", "سایر حساب"),
    ("TC35", "مسکن و قبض منزل", "شخصی رئیس", "سایر حساب"),
    ("TC36", "درمان و سلامت", "شخصی رئیس", "سایر حساب"),
    ("TC37", "تحصیل و آموزش", "شخصی رئیس", "سایر حساب"),
    ("TC38", "سفر و تفریح", "شخصی رئیس", "سایر حساب"),
    ("TC39", "هدیه و نذورات", "شخصی رئیس", "سایر حساب"),
    ("TC40", "دریافتی شخصی", "شخصی رئیس", "سایر حساب"),
    ("TC41", "واریز از حساب شرکت به شخصی", "شخصی رئیس", "سایر حساب"),
    ("TC42", "برگشت وجه شخصی به شرکت", "شخصی رئیس", "سایر حساب"),
]

TXN_OPS = [
    # title, direction, party mode, fixed party, note  (the "nature" of the payment)
    ("واریز نقدی به حساب", "IN", "FIXED", "X003", "پول نقد از صندوق به بانک می‌رود"),
    ("دریافت از مشتری", "IN", "CODE", "", "واریز مشتری به حساب شرکت — کد حساب را بزنید"),
    ("واریز بارفروشان", "IN", "CODE", "", "واریزی بارفروش بابت بار تحویلی"),
    ("پاس چک مشتریان (واریز به حساب ما)", "IN", "CODE", "", "چک دریافتی از مشتری وصول و به حساب ما واریز شد"),
    ("انتقال دریافتی از حساب دیگر", "IN", "CODE", "", "انتقال بانکی ورودی"),
    ("واریز پز باسکول (۲۰۰۶)", "IN", "FIXED", "X011", "سرجمع پز روز قبل؛ حساب واسط X011 تسویه می‌شود"),
    ("دریافت هزینه باسکول دیگران", "IN", "PICK", "", "هزینه باسکول دیگران؛ شخص را از ستون آخر انتخاب کنید"),
    ("دریافت تسهیلات", "IN", "CODE", "", "وام/تسهیلات بانکی"),
    ("سود سپرده", "IN", "FIXED", "X015", "سود سپرده بانکی"),
    ("سایر دریافت", "IN", "CODE", "", ""),
    ("پرداخت به باغدار", "OUT", "CODE", "", "پرداخت بابت خرید مرکبات"),
    ("پرداخت به مشتری", "OUT", "CODE", "", "پرداخت/حواله شرکت به طرف حساب"),
    ("پاس چک ما (چک پرداختی شرکت)", "OUT", "FIXED", "X010", "چک صادرشده توسط شرکت پاس می‌شود"),
    ("برداشت نقدی از حساب", "OUT", "FIXED", "X003", "برداشت نقدی برای صندوق"),
    ("انتقال پرداختی به حساب دیگر", "OUT", "CODE", "", "انتقال بانکی خروجی"),
    ("کارمزد و هزینه بانکی", "OUT", "FIXED", "X012", "کارمزد، آبونمان و هزینه‌های بانکی"),
    ("پرداخت هزینه", "OUT", "CODE", "", "هزینه‌ها: حمل، دستمزد، قبض، تعمیر…"),
    ("پرداخت انعام مصطفی", "OUT", "FIXED", "X001", "بدهی انعام مصطفی تسویه می‌شود"),
    ("پرداخت به باسکول دیگران", "OUT", "PICK", "", "تسویه با شخص باسکول دیگران"),
    ("بازپرداخت تسهیلات", "OUT", "CODE", "", "قسط یا اصل وام"),
    ("پاس چک مشتری برگشتی", "OUT", "CODE", "", "چک مشتری برگشت خورد و از حساب کسر شد"),
    ("سایر پرداخت", "OUT", "CODE", "", ""),
    ("پرداخت به حساب شخصی رئیس", "OUT", "FIXED", "X021",
     "پول از حساب شرکت به مصرف شخصی رئیس می‌رود؛ هزینه شرکت نیست"),
    ("دریافت از حساب شخصی رئیس", "IN", "FIXED", "X021",
     "برگشت پول شخصی رئیس به حساب شرکت"),
    ("انتقال داخلی بین بانک‌ها", "TRF", "NONE", "", "جابه‌جایی بین بانک‌های خود شرکت؛ فقط بانک مقصد را هم انتخاب کنید"),
    ("تسویه دریافت (واسط)", "IN", "FIXED", "X011", "واریزی‌ای که قبلاً در شیت پرداخت (Pay) ثبت شده؛ واسط X011 را صفر می‌کند"),
    ("تسویه پرداخت (واسط)", "OUT", "FIXED", "X011", "پرداختی که قبلاً در شیت پرداخت (Pay) ثبت شده؛ واسط X011 را صفر می‌کند"),
]

TXN_COLS = [
    ("ردیف", "No.", 6,
     "شماره خودکار ردیف؛ با پر شدن تاریخ ساخته می‌شود."),
    ("نوع\nواریز / برداشت", "Type", 11,
     "خودکار از ستون «جزئیات عملیات» ساخته می‌شود: واریز = پول به حساب آمد، "
     "برداشت = پول از حساب رفت، انتقال = بین دو بانک خودمان. رنگ ردیف هم از همین می‌آید."),
    ("کد حساب", "Account code", 13,
     "کد را از کشو انتخاب کنید؛ نام و گروه خودکار می‌آید. اگر فقط نام را می‌دانید، "
     "در شیت Home کادر «جستجوی حساب» را ببینید: نام را انتخاب کنید و کدش را همان‌جا نشان می‌دهد."),
    ("نام حساب", "Account name", 24,
     "خودکار از فایل مادر. اگر «کد ناشناخته» دیدید یعنی این کد در 00_Master تعریف نشده است."),
    ("جزئیات عملیات", "Operation detail", 30,
     "ماهیت پرداخت: پاس چک ما، پاس چک مشتریان، واریز نقدی، برداشت، کارمزد، انتقال داخلی… "
     "همین ستون تعیین می‌کند واریز است یا برداشت و پول از کدام حساب طرف حساب می‌خورد."),
    ("طبقه‌بندی", "Category", 26,
     "این پول برای چه کاری بود؟ خرید مرکبات، فروش مرکبات، قبض، حقوق، تعمیر و نگهداری، "
     "یا هزینه شخصی رئیس. فهرستش در فایل مادر (شیت Cls) تعریف شده و قابل افزودن است."),
    ("دامنه", "Scope", 12,
     "خودکار: «شرکتی» یا «شخصی رئیس». پول شخصی رئیس از حساب شرکت می‌گذرد ولی هزینه شرکت نیست؛ "
     "در شیت Personal جدا گزارش می‌شود."),
    ("تاریخ", "Date", 11,
     "تاریخ شمسی به شکل 1405/07/10."),
    ("سند / پیگیری", "Doc / Ref", 13,
     "شماره سند داخلی یا شماره پیگیری حواله/فیش."),
    ("بانک عامل", "Bank", 20,
     "بانک شرکت که پول از آن خارج یا به آن وارد می‌شود (برای انتقال داخلی: بانک مبدأ)."),
    ("بانک مقصد", "Destination", 16,
     "فقط در «انتقال داخلی» پر می‌شود: بانکی که پول به آن واریز می‌شود."),
    ("روش", "Method", 12,
     "نقد، چک، انتقال بانکی، کارت/پوز یا حواله."),
    ("مبلغ (ریال)", "Amount", 15,
     "مبلغ کامل تراکنش؛ بدهکار و بستانکار خودکار از آن ساخته می‌شود."),
    ("شرح", "Description", 26,
     "توضیح کوتاه؛ در دفتر بانک و گزارش‌ها دیده می‌شود."),
    ("بدهکار بانک", "Bank Dr", 12, "خودکار: پول وارد حساب بانک شده است."),
    ("بستانکار بانک", "Bank Cr", 12, "خودکار: پول از حساب بانک خارج شده است."),
    ("شخص باسکول دیگران", "Other scale", 20,
     "فقط برای عملیات باسکول دیگران: شخص را انتخاب کنید تا بدهی‌اش در شیت Payable ردیابی شود."),
    ("وضعیت", "Status", 11,
     "✓ یعنی ردیف کامل است؛ ⚠ یعنی چیزی جا افتاده (مبلغ، بانک، طبقه‌بندی یا کد حساب)."),
    ("یادداشت", "Note", 20, "هر توضیح اضافه."),
]


def sheet_txn(wb, k):
    ws = wb.add_worksheet("Txn")
    last = len(TXN_COLS) - 1          # 18 = S
    k.setup(ws, last, 12 + NR + 2, landscape=True, repeat=(11, 11), freeze=(12, 4))
    k.brand(ws, last, "دفتر تراکنش‌های بانکی — واریز، برداشت و ردیابی",
            "از راست: ردیف ← نوع خودکار ← کد حساب ← نام خودکار ← جزئیات عملیات ← طبقه‌بندی  ·  "
            "رنگ ردیف با ماهیت پرداخت تغییر می‌کند", icon="bank")
    k.kpis(ws, 6, [
        ("تعداد تراکنش", f'=COUNTA($H$13:$H${12+NR})', INT, FOREST, "list"),
        ("جمع واریز", f'=SUM($O$13:$O${12+NR})', INT, GREEN, "transfer"),
        ("جمع برداشت", f'=SUM($P$13:$P${12+NR})', INT, RED, "transfer"),
        ("خالص جریان", f'=SUM($O$13:$O${12+NR})-SUM($P$13:$P${12+NR})', INT_RED, BLUE, "money"),
        ("ردیف‌های ناقص", f'=COUNTIF($R$13:$R${12+NR},"⚠*")', INT, TERRA, "warn"),
    ], last)

    # colour legend
    legend = [("واریز — پول وارد حساب شد", MINT, GREEN),
              ("برداشت — پول از حساب رفت", BLUSH, TERRA),
              ("انتقال داخلی — فقط جابه‌جایی بین بانک‌های خودمان", LAVENDER, BLUE)]
    c = 0
    for text, bg, fg in legend:
        span = max(2, (last + 1) // 3)
        ws.merge_range(8, c, 8, min(c + span - 1, last), "     " + text,
                       k.f(font_name=FA, font_size=9, bold=True, font_color=fg,
                           bg_color=bg, align="right", valign="vcenter",
                           border=1, border_color=WHITE))
        c += span
    ws.set_row(8, 24)

    # ---- name → code finder (row 10): type/choose the account name, get the code
    ws.merge_range(9, 1, 9, 3, "   کد حساب را نمی‌دانید؟ نام حساب را از کشو انتخاب کنید:",
                   k.f(font_name=FA, font_size=9, bold=True, font_color=FOREST,
                       bg_color=MINT, align="right", valign="vcenter",
                       border=1, border_color=SAGE))
    ws.write_blank(9, 4, None, k.f(font_name=FA, font_size=9, bold=True, bg_color="#FFF4D6",
                                   border=1, border_color=GOLD_L, align="right",
                                   valign="vcenter"))
    ws.data_validation(9, 4, 9, 4, {"validate": "list", "source": "=ListAcctNames",
                                    "input_title": "پیدا کردن کد از نام",
                                    "input_message": "نام حساب را انتخاب کنید؛ کدش کنارش نوشته می‌شود.",
                                    "error_title": "نام حساب", "error_message": "از فهرست انتخاب کنید."})
    ws.merge_range(9, 5, 9, 7,
                   '=IF($E$10="","کد: —","کد: "&IFERROR(INDEX(Sync!$H$10:$H$300,'
                   'MATCH($E$10,Sync!$I$10:$I$300,0)),"پیدا نشد"))',
                   k.f(font_name=FA, font_size=9, bold=True, font_color=BLUE,
                       bg_color="#EEF3F8", align="center", valign="vcenter",
                       border=1, border_color=WHITE))
    ws.set_row(9, 22)

    k.section(ws, 10, last, "۰۱", "دفتر تراکنش‌ها",
              "ستون‌های طلایی را پر کنید؛ ستون‌های خاکستری خودکارند")
    k.thead(ws, 11, TXN_COLS, FOREST)
    ws.set_row(11, 40)

    f_in = k.f(font_name=EN, font_size=9, bg_color="#FFF4D6", border=1,
               border_color=GOLD_L, align="center", valign="vcenter")
    f_in_t = k.f(font_name=FA, font_size=9.5, bg_color="#FFF4D6", border=1,
                 border_color=GOLD_L, align="right", valign="vcenter")
    f_in_n = k.f(font_name=EN, font_size=9.5, bg_color="#FFF4D6", border=1,
                 border_color=GOLD_L, align="center", valign="vcenter", num_format=INT)
    f_auto = k.f(font_name=EN, font_size=8.5, bg_color="#F6F5F0", border=1,
                 border_color=LINE_SOFT, align="center", valign="vcenter")
    f_auto_t = k.f(font_name=FA, font_size=9, bg_color="#F6F5F0", border=1,
                   border_color=LINE_SOFT, align="right", valign="vcenter")
    f_auto_n = k.f(font_name=EN, font_size=9, bg_color="#F6F5F0", border=1,
                   border_color=LINE_SOFT, align="center", valign="vcenter", num_format=INT)
    f_type = k.f(font_name=FA, font_size=9.5, bold=True, bg_color="#F6F5F0", border=1,
                 border_color=LINE_SOFT, align="center", valign="vcenter")
    f_scope = k.f(font_name=FA, font_size=8.5, bg_color="#F6F5F0", border=1,
                  border_color=LINE_SOFT, align="center", valign="vcenter")
    f_status = k.f(font_name=FA, font_size=9, bold=True, bg_color=CREAM, border=1,
                   border_color=LINE_SOFT, align="center", valign="vcenter")

    for i in range(NR):
        r = 12 + i
        x = r + 1
        ws.set_row(r, 21)
        ws.write_formula(r, 0, f'=IF($H{x}="","",COUNTA($H$13:$H{x}))', f_auto)
        ws.write_formula(r, 1,
                         f'=IF($E{x}="","",IF($T{x}="TRF","انتقال",IF($T{x}="IN","واریز","برداشت")))',
                         f_type)
        ws.write_blank(r, 2, None, f_in)                    # C code (input)
        ws.write_formula(r, 3,
                         f'=IF($C{x}="","",IFERROR(VLOOKUP($C{x},AcctTbl,2,0),"کد ناشناخته"))',
                         f_auto_t)
        ws.write_blank(r, 4, None, f_in_t)                  # E operation detail
        ws.write_blank(r, 5, None, f_in_t)                  # F classification
        ws.write_formula(r, 6,
                         f'=IF($F{x}="","",IFERROR(VLOOKUP($F{x},ClassTbl,3,0),"شرکتی"))', f_scope)
        ws.write_blank(r, 7, None, f_in)                    # H date
        ws.write_blank(r, 8, None, f_in)                    # I doc
        ws.write_blank(r, 9, None, f_in_t)                  # J bank
        ws.write_blank(r, 10, None, f_in_t)                 # K dest
        ws.write_blank(r, 11, None, f_in)                   # L method
        ws.write_blank(r, 12, None, f_in_n)                 # M amount
        ws.write_blank(r, 13, None, f_in_t)                 # N desc
        ws.write_formula(r, 14, f'=IF($H{x}="",0,IF($T{x}="IN",N($M{x}),0))', f_auto_n)
        ws.write_formula(r, 15,
                         f'=IF($H{x}="",0,IF(OR($T{x}="OUT",$T{x}="TRF"),N($M{x}),0))', f_auto_n)
        ws.write_blank(r, 16, None, f_in_t)                 # Q other scale person
        ws.write_formula(
            r, 17,
            f'=IF($H{x}="","",IF($M{x}="","⚠ مبلغ",IF($E{x}="","⚠ عملیات",'
            f'IF($J{x}="","⚠ بانک",IF(AND($T{x}="TRF",$K{x}=""),"⚠ بانک مقصد",'
            f'IF($F{x}="","⚠ طبقه‌بندی",IF(AND($U{x}<>"NONE",$U{x}<>"",$Y{x}=""),'
            f'"⚠ کد حساب","✓")))))))', f_status)
        ws.write_blank(r, 18, None, f_in_t)                 # S note

        # ---------------- hidden helpers T..AC
        ws.write_formula(r, 19, f'=IFERROR(VLOOKUP($E{x},BankOpTbl,2,0),"")', f_auto)   # T dir
        ws.write_formula(r, 20, f'=IFERROR(VLOOKUP($E{x},BankOpTbl,3,0),"")', f_auto)   # U mode
        ws.write_formula(r, 21, f'=IFERROR(VLOOKUP($E{x},BankOpTbl,4,0),"")', f_auto)   # V fixed
        ws.write_formula(r, 22,
                         f'=IF($J{x}="","",IFERROR(VLOOKUP($J{x},BankTbl,4,0),""))', f_auto)  # W
        ws.write_formula(r, 23,
                         f'=IF($K{x}="","",IFERROR(VLOOKUP($K{x},BankTbl,4,0),""))', f_auto)  # X
        ws.write_formula(
            r, 24,
            f'=IF($H{x}="","",IF($U{x}="PICK",LEFT($Q{x},5),IF($U{x}="FIXED",$V{x},'
            f'IF($U{x}="CODE",$C{x},""))))', f_auto)                                     # Y party
        ws.write_formula(r, 25,
                         f'=IF($F{x}="","",IFERROR(VLOOKUP($F{x},ClassTbl,3,0),"شرکتی"))', f_auto)
        ws.write_formula(r, 26,
                         f'=IF($F{x}="","",IFERROR(VLOOKUP($F{x},ClassTbl,4,0),""))', f_auto)
        ws.write_formula(r, 27,
                         f'=IF($H{x}="",0,IF(OR($T{x}="OUT",$T{x}="TRF"),N($M{x}),0))', f_auto_n)
        ws.write_formula(r, 28,
                         f'=IF($H{x}="",0,IF($T{x}="IN",N($M{x}),0))', f_auto_n)
        ws.write_formula(r, 29,
                         f'=IF($H{x}="","",IFERROR(VALUE(MID($H{x},6,2)),0))', f_auto)

    for c_, w in ((19, 6), (20, 8), (21, 8), (22, 9), (23, 9), (24, 10), (25, 12),
                  (26, 10), (27, 12), (28, 12), (29, 6)):
        ws.set_column(c_, c_, w)
        ws.set_column(c_, c_, w, None, {"hidden": True})

    # demo rows: (date, doc, op, bank, dest, method, code, desc, ref, amount, person, class)
    for i, row in enumerate(DEMO):
        r = 12 + i
        date, doc, op, bank, dest, method, code, desc, ref, amount = row[:10]
        person = row[10] if len(row) > 10 else ""
        cls = row[11] if len(row) > 11 else ""
        ws.write_string(r, 7, date, f_in)
        ws.write_string(r, 8, doc, f_in)
        ws.write_string(r, 4, op, f_in_t)
        ws.write_string(r, 9, bank, f_in_t)
        if dest:
            ws.write_string(r, 10, dest, f_in_t)
        ws.write_string(r, 11, method, f_in)
        if code:
            ws.write_string(r, 2, code, f_in)
        if cls:
            ws.write_string(r, 5, cls, f_in_t)
        ws.write_string(r, 13, desc, f_in_t)
        ws.write_number(r, 12, amount, f_in_n)
        if person:
            ws.write_string(r, 16, person, f_in_t)
        if ref:
            ws.write_string(r, 18, f"پیگیری: {ref}", f_in_t)

    total = 12 + NR
    ws.merge_range(total, 0, total, 13, "جمع کل", 
                   k.f(font_name=FA, font_size=10.5, bold=True, font_color=WHITE,
                       bg_color=GOLD, align="center", valign="vcenter", border=1,
                       border_color=GOLD))
    for c_ in (14, 15, 27, 28):
        ws.write_formula(total, c_, f"=SUM({chr(65+c_)}{12}:{chr(65+c_)}{total-1})",
                         k.f(font_name=EN, font_size=10, bold=True, font_color=WHITE,
                            bg_color=GOLD, align="center", valign="vcenter", border=1,
                            border_color=GOLD, num_format=INT))
    ws.write_formula(total, 17,
                     f'=IF(COUNTIF($R$13:$R${total},"⚠*")=0,"✓ همه ردیف‌ها کامل","⚠ بررسی")',
                     k.f(font_name=FA, font_size=10, bold=True, font_color=WHITE,
                        bg_color=GOLD, align="center", valign="vcenter", border=1,
                        border_color=GOLD))
    ws.set_row(total, 24)

    # validations
    ws.data_validation(12, 7, 12 + NR - 1, 7, {
        "validate": "custom",
        "value": '=OR($H13="",AND(LEN($H13)=10,MID($H13,5,1)="/",MID($H13,8,1)="/"))',
        "error_title": "تاریخ شمسی",
        "error_message": "تاریخ را به شکل 1405/07/10 وارد کنید.",
        "error_type": "stop"})
    ws.data_validation(12, 4, 12 + NR - 1, 4, {"validate": "list", "source": "=ListTxnOps",
                                               "error_type": "stop"})
    ws.data_validation(12, 5, 12 + NR - 1, 5, {"validate": "list", "source": "=ListClasses",
                                               "error_type": "warning"})
    ws.data_validation(12, 2, 12 + NR - 1, 2, {"validate": "list", "source": "=ListAllCodes",
                                               "error_type": "warning"})
    ws.data_validation(12, 9, 12 + NR - 1, 9, {"validate": "list", "source": "=ListBanks",
                                               "error_type": "stop"})
    ws.data_validation(12, 10, 12 + NR - 1, 10, {"validate": "list", "source": "=ListBanks",
                                                 "error_type": "warning"})
    ws.data_validation(12, 11, 12 + NR - 1, 11, {"validate": "list", "source": "=ListPayMethods",
                                                 "error_type": "warning"})
    ws.data_validation(12, 16, 12 + NR - 1, 16, {"validate": "list",
                                                 "source": "=ListOtherScales",
                                                 "error_type": "warning"})

    # colour by direction + type chip
    ws.conditional_format(11, 0, 12 + NR - 1, last, {
        "type": "formula", "criteria": '=$T12="IN"', "format": k.f(bg_color="#F2FAF4")})
    ws.conditional_format(11, 0, 12 + NR - 1, last, {
        "type": "formula", "criteria": '=$T12="OUT"', "format": k.f(bg_color="#FDF3F0")})
    ws.conditional_format(11, 0, 12 + NR - 1, last, {
        "type": "formula", "criteria": '=$T12="TRF"', "format": k.f(bg_color="#F4F2FA")})
    ws.conditional_format(11, 1, 12 + NR - 1, 1, {
        "type": "formula", "criteria": '=$T12="IN"',
        "format": k.f(bg_color=MINT, font_color=GREEN, bold=True)})
    ws.conditional_format(11, 1, 12 + NR - 1, 1, {
        "type": "formula", "criteria": '=$T12="OUT"',
        "format": k.f(bg_color="#F8DDD3", font_color=TERRA, bold=True)})
    ws.conditional_format(11, 1, 12 + NR - 1, 1, {
        "type": "formula", "criteria": '=$T12="TRF"',
        "format": k.f(bg_color=LAVENDER, font_color=BLUE, bold=True)})
    ws.conditional_format(11, 6, 12 + NR - 1, 6, {
        "type": "text", "criteria": "containing", "value": "شخصی",
        "format": k.f(bg_color="#FFF1D6", font_color="#9A6B15", bold=True)})
    ws.conditional_format(11, 17, 12 + NR - 1, 17, {
        "type": "text", "criteria": "containing", "value": "⚠",
        "format": k.f(bg_color="#FBE3D8", font_color=TERRA, bold=True)})
    ws.conditional_format(11, 17, 12 + NR - 1, 17, {
        "type": "text", "criteria": "containing", "value": "✓",
        "format": k.f(bg_color=MINT, font_color=GREEN, bold=True)})
    ws.autofilter(10, 0, 12 + NR - 1, last)
    return ws


# =====================================================================  Post
def sheet_post(wb, k):
    ws = wb.add_worksheet("Post")
    ws.hide()
    k.setup(ws, 15, 12 + PN, landscape=True, freeze=(12, 1))
    cols = ["تاریخ", "سند", "منبع", "کد حساب", "نام حساب", "گروه", "عملیات", "بانک",
            "بدهکار", "بستانکار", "ماه", "شرح", "", "بانک؟", "دامنه", "طبقه‌بندی"]
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
                (p, "O", "P", f"Txn!$W{x}",
                 f'IF(Txn!$N{x}="","",Txn!$N{x})'),
                (p + 1, "AB", "AC",
                 f'IF(Txn!$T{x}="TRF",Txn!$X{x},Txn!$Y{x})',
                 f'IF(Txn!$N{x}="","","طرف مقابل — "&Txn!$N{x})'))):
            ws.write_formula(row, 0, f'=IF(Txn!$H{x}="","",Txn!$H{x})', f_t)
            ws.write_formula(row, 1, f'=IF(Txn!$H{x}="","",Txn!$I{x})', f_t)
            ws.write_formula(row, 2, f'=IF(Txn!$H{x}="","",$P{row+1}&" · "&$O{row+1})', f_free)
            ws.write_formula(row, 3, f'=IF(Txn!$H{x}="","",IF({acct}="","",{acct}))', f_t)
            ws.write_formula(row, 4, f'=IF($D{row+1}="","",IFERROR(VLOOKUP($D{row+1},AcctTbl,2,0),""))', f_free)
            ws.write_formula(row, 5, f'=IF($D{row+1}="","",IFERROR(VLOOKUP($D{row+1},AcctTbl,3,0),""))', f_free)
            ws.write_formula(row, 6, f'=IF(Txn!$H{x}="","",Txn!$E{x})', f_free)
            ws.write_formula(row, 7, f'=IF(Txn!$H{x}="","",Txn!$J{x})', f_t)
            ws.write_formula(row, 8, f'=IF(Txn!$H{x}="",0,Txn!${dr}{x})', f_n)
            ws.write_formula(row, 9, f'=IF(Txn!$H{x}="",0,Txn!${cr}{x})', f_n)
            ws.write_formula(row, 10,
                             f'=IF($A{row+1}="","",IFERROR(VALUE(MID($A{row+1},6,2)),""))',
                             k.f(font_name=EN, font_size=8, num_format="00", align="center",
                                 bg_color="#F6F5F0"))
            ws.write_formula(row, 11, f'={desc}', f_free)
            ws.write_formula(row, 13,
                             f'=IF($D{row+1}="",0,COUNTIF(BankAccts,$D{row+1}))',
                             k.f(font_name=EN, font_size=8, align="center",
                                 bg_color="#F6F5F0"))
            ws.write_formula(row, 14, f'=IF(Txn!$H{x}="","",Txn!$Z{x})', f_free)
            ws.write_formula(row, 15, f'=IF(Txn!$H{x}="","",Txn!$F{x})', f_free)
    for c_, w in enumerate([11, 11, 30, 11, 26, 14, 24, 11, 14, 14, 7, 30, 2, 7, 12, 22]):
        ws.set_column(c_, c_, w)
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
    nops = len(TXN_OPS)
    k.setup(ws, last, 40 + nops, landscape=True, repeat=(7, 7), freeze=(10, 1))
    k.brand(ws, last, "تحلیل تراکنش‌ها",
            "بر پایه نوع عملیات، روش پرداخت و طرف حساب  ·  همه از دفتر تراکنش‌ها", icon="chart")
    k.kpis(ws, 6, [
        ("جمع واریز", f"=SUM(Txn!$O$13:$O${12+NR})", INT, GREEN, "transfer"),
        ("جمع برداشت", f"=SUM(Txn!$P$13:$P${12+NR})", INT, RED, "transfer"),
        ("خالص جریان", f"=SUM(Txn!$O$13:$O${12+NR})-SUM(Txn!$P$13:$P${12+NR})", INT_RED, FOREST,
         "money"),
        ("سهم شخصی از گردش", f'=IFERROR(SUMIFS(Txn!$M$13:$M${12+NR},Txn!$Z$13:$Z${12+NR},'
                             f'"شخصی رئیس")/MAX(1,SUM(Txn!$M$13:$M${12+NR})),0)',
         PCT, GOLD, "people"),
    ], last)

    head_f = k.f(font_name=FA, font_size=9.5, bold=True, font_color=WHITE, bg_color=FOREST,
                 border=1, border_color=FOREST, align="center", valign="vcenter",
                 text_wrap=True)
    en_f = k.f(font_name=EN, font_size=6.8, font_color=SAGE_L, bg_color=FOREST, align="center")
    lab_f = k.f(font_name=FA, font_size=9, bg_color=SAND_2, border=1, border_color=LINE,
                align="right")
    in_f = k.f(font_name=EN, font_size=9, bg_color=MINT, border=1, border_color=LINE_SOFT,
               align="center", num_format=INT)
    out_f = k.f(font_name=EN, font_size=9, bg_color=BLUSH, border=1, border_color=LINE_SOFT,
                align="center", num_format=INT)
    net_f = k.f(font_name=EN, font_size=9, bold=True, bg_color=CREAM, border=1,
                border_color=LINE_SOFT, align="center", num_format=INT_RED)
    cnt_f = k.f(font_name=EN, font_size=9, bg_color=WHITE, border=1, border_color=LINE_SOFT,
                align="center", num_format=INT)
    pct_f = k.f(font_name=EN, font_size=8.5, bg_color=WHITE, border=1, border_color=LINE_SOFT,
                align="center", num_format=PCT)
    bar_f = k.f(font_name=EN, font_size=7, font_color=GOLD, bg_color=WHITE, align="right")

    # ---------------------------------------------------------------- ۰۱ ops
    k.section(ws, 8, last, "۰۱", "تحلیل بر اساس جزئیات عملیات", "کدام نوع عملیات بیشترین گردش را دارد؟")
    heads = [("جزئیات عملیات", "Operation", 30), ("تعداد", "Count", 9), ("واریز", "In", 15),
             ("برداشت", "Out", 15), ("خالص", "Net", 15), ("سهم برداشت", "Share", 10),
             ("نمودار", "Bar", 8)]
    for i, (fa, en, w) in enumerate(heads):
        ws.set_column(i, i, w)
        ws.write(9, i, fa, head_f)
        ws.write(10, i, en, en_f)
    r0 = 11
    for i, (title, *_rest) in enumerate(TXN_OPS):
        r = r0 + i
        ws.write_string(r, 0, title, lab_f)
        ws.write_formula(r, 1, f'=COUNTIF(Txn!$E$13:$E${12+NR},$A{r+1})', cnt_f)
        ws.write_formula(r, 2, f'=SUMIFS(Txn!$M$13:$M${12+NR},Txn!$E$13:$E${12+NR},$A{r+1},'
                               f'Txn!$T$13:$T${12+NR},"IN")', in_f)
        ws.write_formula(r, 3, f'=SUMIFS(Txn!$M$13:$M${12+NR},Txn!$E$13:$E${12+NR},$A{r+1},'
                               f'Txn!$T$13:$T${12+NR},"OUT")', out_f)
        ws.write_formula(r, 4, f'=$C{r+1}-$D{r+1}', net_f)
        ws.write_formula(r, 5, f'=IFERROR($D{r+1}/MAX(1,SUM($D${r0+1}:$D${r0+nops})),0)', pct_f)
        ws.write_formula(r, 6, f'=REPT("█",MIN(16,ROUND(MAX($C{r+1},$D{r+1})/'
                               f'MAX(1,MAX($C${r0+1}:$C${r0+nops}),MAX($D${r0+1}:$D${r0+nops}))*16,0)))',
                         bar_f)
    ch = wb.add_chart({"type": "doughnut"})
    ch.add_series({
        "name": "سهم برداشت",
        "categories": ["Anal", r0, 0, r0 + nops - 1, 0],
        "values": ["Anal", r0, 3, r0 + nops - 1, 3],
        "points": [{"fill": {"color": c}} for c in
                   (FOREST, SAGE, GOLD, TERRA, BLUE, AMBER, "#8E6C88", "#4F7C6C",
                    "#B08968", "#7D8CA3", "#C9ADA7", "#8CB369")],
        "data_labels": {"percentage": True, "font": {"size": 7}},
    })
    ch.set_title({"name": "سهم هر عملیات از برداشت‌ها", "name_font": {"name": FA, "size": 10}})
    ch.set_legend({"position": "right", "font": {"name": FA, "size": 7}})
    ch.set_size({"width": 400, "height": 300})
    ws.insert_chart("I11", ch, {"x_offset": 6, "y_offset": 4})

    # ------------------------------------------------------------ ۰۲ methods
    m1 = r0 + nops + 2
    k.section(ws, m1, last, "۰۲", "تحلیل بر اساس روش پرداخت", "نقد، چک، انتقال، کارت/پوز، حواله")
    mh = ["روش", "تعداد", "واریز", "برداشت", "خالص", "سهم"]
    for i, h in enumerate(mh):
        ws.write(m1 + 1, i, h, k.f(font_name=FA, font_size=9, bold=True, font_color=WHITE,
                                   bg_color=FOREST_2, border=1, border_color=FOREST_2,
                                   align="center"))
    for i, m in enumerate(["نقد", "چک", "انتقال بانکی", "کارت/پوز", "حواله"]):
        r = m1 + 2 + i
        ws.write_string(r, 0, m, lab_f)
        ws.write_formula(r, 1, f'=COUNTIFS(Txn!$L$13:$L${12+NR},$A{r+1})', cnt_f)
        ws.write_formula(r, 2, f'=SUMIFS(Txn!$O$13:$O${12+NR},Txn!$L$13:$L${12+NR},$A{r+1})', in_f)
        ws.write_formula(r, 3, f'=SUMIFS(Txn!$P$13:$P${12+NR},Txn!$L$13:$L${12+NR},$A{r+1})', out_f)
        ws.write_formula(r, 4, f'=$C{r+1}-$D{r+1}', net_f)
        ws.write_formula(r, 5, f'=IFERROR($D{r+1}/MAX(1,SUM($D${m1+2}:$D${m1+6})),0)', pct_f)
    ch2 = wb.add_chart({"type": "pie"})
    ch2.add_series({"name": "روش پرداخت", "categories": ["Anal", m1 + 2, 0, m1 + 6, 0],
                    "values": ["Anal", m1 + 2, 3, m1 + 6, 3],
                    "points": [{"fill": {"color": c}} for c in (SAGE, GOLD, FOREST, TERRA, BLUE)],
                    "data_labels": {"percentage": True, "font": {"size": 8}}})
    ch2.set_title({"name": "ترکیب روش‌های پرداخت", "name_font": {"name": FA, "size": 10}})
    ch2.set_legend({"position": "right", "font": {"name": FA, "size": 7}})
    ch2.set_size({"width": 340, "height": 230})
    ws.insert_chart(f"I{m1+2}", ch2, {"x_offset": 6, "y_offset": 4})

    # ------------------------------------------------------------ ۰۳ parties
    p1 = m1 + 9
    k.section(ws, p1, last, "۰۳", "بزرگ‌ترین طرف حساب‌ها",
              "جمع گردش هر کد حساب در دفتر بانک")
    th = ["کد حساب", "نام", "گروه", "جمع بدهکار", "جمع بستانکار", "خالص"]
    for i, h in enumerate(th):
        ws.write(p1 + 1, i, h, k.f(font_name=FA, font_size=9, bold=True, font_color=WHITE,
                                   bg_color=FOREST, border=1, border_color=FOREST,
                                   align="center"))
    for i in range(8):
        r = p1 + 2 + i
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
        ws.write_formula(r, 3, f'=IF($A{r+1}="","",SUMIF(Post!$D$13:$D${12+PN},$A{r+1},'
                               f'Post!$I$13:$I${12+PN}))',
                         k.f(font_name=EN, font_size=9, bg_color=WHITE, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(r, 4, f'=IF($A{r+1}="","",SUMIF(Post!$D$13:$D${12+PN},$A{r+1},'
                               f'Post!$J$13:$J${12+PN}))',
                         k.f(font_name=EN, font_size=9, bg_color=WHITE, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(r, 5, f'=IF($A{r+1}="","",$D{r+1}-$E{r+1})', net_f)
    note_row = p1 + 12
    k.note(ws, note_row, last,
           "برای دیدن ریز هر طرف حساب، شیت Post را فیلتر کنید یا از فایل گزارش (03) استفاده کنید. "
           "تحلیل طبقه‌بندی و هزینه‌های شخصی رئیس در شیت‌های Class و Personal است.",
           bg=SAND_2, icon="search")
    ws.print_area(0, 0, note_row, last)
    return ws


# =====================================================================  Class
def sheet_class(wb, k, classes):
    """Turnover of every classification, split into company / owner-personal."""
    ws = wb.add_worksheet("Class")
    last = 8
    n = len(classes)
    k.setup(ws, last, 14 + n + 6, landscape=True, repeat=(12, 13), freeze=(14, 3))
    k.brand(ws, last, "طبقه‌بندی تراکنش‌ها — پول برای چه کاری رفت؟",
            "هر تراکنش یک طبقه‌بندی دارد: خرید مرکبات، قبض، حقوق، تعمیر دستگاه‌ها یا هزینه شخصی رئیس",
            icon="list")
    k.kpis(ws, 6, [
        ("جمع گردش شرکتی", f'=SUMIF($B$14:$B${13+n},"شرکتی",$H$14:$H${13+n})', INT, FOREST,
         "money"),
        ("جمع گردش شخصی رئیس", f'=SUMIF($B$14:$B${13+n},"شخصی رئیس",$H$14:$H${13+n})', INT,
         GOLD, "people"),
        ("سهم شخصی", f'=IFERROR($C$8/MAX(1,$A$8+$C$8),0)', PCT, TERRA, "gauge"),
        ("طبقه‌بندی‌های دارای گردش", f'=COUNTIF($F$14:$F${13+n},">0")', INT, BLUE, "list"),
    ], last)
    k.note(ws, 9, last,
           "ستون «طبقه‌بندی» در شیت Txn از همین فهرست می‌آید. برای افزودن طبقه‌بندی جدید: "
           "00_Master ← شیت Cls ← یک ردیف خالی را پر کنید و فایل‌ها را دوباره بسازید.", bg=SAND,
           icon="flag")
    k.section(ws, 11, last, "۰۱", "گردش هر طبقه‌بندی",
              "دامنه = شرکتی یا شخصی رئیس  ·  مانده = واریز − برداشت")
    heads = [("طبقه‌بندی", "Class", 30), ("دامنه", "Scope", 12), ("گروه", "Group", 16),
             ("تعداد", "Count", 9), ("واریز", "In", 15), ("برداشت", "Out", 15),
             ("مانده", "Net", 15), ("سهم برداشت", "Share", 10), ("نمودار", "Bar", 10)]
    for i, (fa, en, w) in enumerate(heads):
        ws.set_column(i, i, w)
        ws.write(12, i, fa, k.f(font_name=FA, font_size=9, bold=True, font_color=WHITE,
                                bg_color=FOREST, border=1, border_color=FOREST,
                                align="center", valign="vcenter", text_wrap=True))
        ws.write(13, i, en, k.f(font_name=EN, font_size=6.5, font_color=SAGE_L,
                                bg_color=FOREST, align="center"))
    lab = k.f(font_name=FA, font_size=9, bg_color=SAND_2, border=1, border_color=LINE,
              align="right")
    sc = k.f(font_name=FA, font_size=8.5, bg_color=WHITE, border=1, border_color=LINE_SOFT,
             align="center")
    grp = k.f(font_name=FA, font_size=8.5, bg_color=WHITE, border=1, border_color=LINE_SOFT,
              align="right")
    cnt = k.f(font_name=EN, font_size=9, bg_color=WHITE, border=1, border_color=LINE_SOFT,
              align="center", num_format=INT)
    in_f = k.f(font_name=EN, font_size=9, bg_color=MINT, border=1, border_color=LINE_SOFT,
               align="center", num_format=INT)
    out_f = k.f(font_name=EN, font_size=9, bg_color=BLUSH, border=1, border_color=LINE_SOFT,
                align="center", num_format=INT)
    net_f = k.f(font_name=EN, font_size=9, bold=True, bg_color=CREAM, border=1,
                border_color=LINE_SOFT, align="center", num_format=INT_RED)
    pct = k.f(font_name=EN, font_size=8.5, bg_color=WHITE, border=1, border_color=LINE_SOFT,
              align="center", num_format=PCT)
    bar = k.f(font_name=EN, font_size=7, font_color=GOLD, bg_color=WHITE, align="right")
    r0 = 14
    for i, (code, title, scope, group) in enumerate(classes):
        r = r0 + i
        ws.write_string(r, 0, title, lab)
        ws.write_string(r, 1, scope, sc)
        ws.write_string(r, 2, group, grp)
        ws.write_formula(r, 3, f'=COUNTIF(Txn!$F$13:$F${12+NR},$A{r+1})', cnt)
        ws.write_formula(r, 4, f'=SUMIFS(Txn!$M$13:$M${12+NR},Txn!$F$13:$F${12+NR},$A{r+1},'
                               f'Txn!$T$13:$T${12+NR},"IN")', in_f)
        ws.write_formula(r, 5, f'=SUMIFS(Txn!$M$13:$M${12+NR},Txn!$F$13:$F${12+NR},$A{r+1},'
                               f'Txn!$T$13:$T${12+NR},"OUT")', out_f)
        ws.write_formula(r, 6, f'=$E{r+1}-$F{r+1}', net_f)
        ws.write_formula(r, 7, f'=IFERROR($F{r+1}/MAX(1,SUM($F${r0+1}:$F${r0+n})),0)', pct)
        ws.write_formula(r, 8, f'=REPT("█",MIN(14,ROUND(MAX($E{r+1},$F{r+1})/'
                               f'MAX(1,MAX($E${r0+1}:$E${r0+n}),MAX($F${r0+1}:$F${r0+n}))*14,0)))',
                         bar)
    tot = r0 + n
    ws.merge_range(tot, 0, tot, 2, "جمع کل", k.f(font_name=FA, font_size=10, bold=True,
                                                 font_color=WHITE, bg_color=GOLD, border=1))
    for c_ in (3, 4, 5, 6):
        ws.write_formula(tot, c_, f"=SUM({chr(65+c_)}{r0+1}:{chr(65+c_)}{tot})",
                         k.f(font_name=EN, font_size=9.5, bold=True, font_color=WHITE,
                             bg_color=GOLD, border=1, num_format=INT))
    ch = wb.add_chart({"type": "bar"})
    ch.add_series({"name": "برداشت", "categories": ["Class", r0, 0, r0 + 11, 0],
                   "values": ["Class", r0, 5, r0 + 11, 5], "fill": {"color": TERRA},
                   "gap": 40})
    ch.add_series({"name": "واریز", "categories": ["Class", r0, 0, r0 + 11, 0],
                   "values": ["Class", r0, 4, r0 + 11, 4], "fill": {"color": SAGE},
                   "gap": 40})
    ch.set_title({"name": "ده طبقه‌بندی نخست (واریز و برداشت)",
                  "name_font": {"name": FA, "size": 11}})
    ch.set_legend({"position": "bottom", "font": {"name": FA, "size": 8}})
    ch.set_size({"width": 560, "height": 280})
    ws.insert_chart(f"K{r0}", ch, {"x_offset": 8, "y_offset": 6})
    ws.print_area(0, 0, tot, last)
    return ws


# ==================================================================  Personal
def sheet_personal(wb, k, classes):
    """Owner's private money that runs through the company bank accounts."""
    ws = wb.add_worksheet("Personal")
    last = 7
    pers = [c for c in classes if c[2] != "شرکتی"]
    n = max(len(pers), 1)
    k.setup(ws, last, 52, landscape=False, repeat=(12, 12), freeze=(13, 2))
    k.brand(ws, last, "پول شخصی رئیس شرکت — جدا از هزینه‌های شرکت",
            "این مبالغ از حساب شرکت می‌گذرند ولی هزینه شرکت نیستند؛ به حساب جاری رئیس می‌روند",
            icon="people")
    k.kpis(ws, 6, [
        ("برداشت شخصی (از شرکت)", f'=SUM($E$14:$E${13+n})', INT, TERRA, "transfer"),
        ("برگشت پول شخصی به شرکت", f'=SUM($D$14:$D${13+n})', INT, GREEN, "transfer"),
        ("مانده بدهی رئیس به شرکت", f'=SUM($E$14:$E${13+n})-SUM($D$14:$D${13+n})', INT_RED,
         FOREST, "gauge"),
        ("تعداد تراکنش شخصی", f'=SUM($C$14:$C${13+n})', INT, GOLD, "list"),
    ], last)
    k.note(ws, 9, last,
           "قاعده: هر پولی که برای مصرف شخصی رئیس (خانه، خودرو، وسایل، درمان، سفر) از حساب شرکت "
           "بیرون می‌رود، در ستون «طبقه‌بندی» شیت Txn با یکی از طبقه‌بندی‌های «شخصی رئیس» ثبت می‌شود. "
           "این مبلغ در دفتر به حساب جاری رئیس (X021) می‌نشیند، نه به هزینه شرکت؛ پس سود شرکت را کم نمی‌کند "
           "و بدهی رئیس به شرکت را نشان می‌دهد.", bg="#FFF1D6", icon="flag", h=56)
    k.section(ws, 11, last, "۰۱", "ریز مصرف شخصی",
              "هر ردیف یک نوع هزینه شخصی  ·  ستون آخر نمای نسبی")
    heads = [("طبقه‌بندی شخصی", "Class", 30), ("تعداد", "Count", 9), ("واریز/برگشت", "In", 15),
             ("برداشت", "Out", 15), ("مانده", "Net", 15), ("نمودار", "Bar", 12)]
    for i, (fa, en, w) in enumerate(heads):
        ws.set_column(i, i, w)
        ws.write(12, i, fa, k.f(font_name=FA, font_size=9, bold=True, font_color=WHITE,
                                bg_color=TERRA, border=1, border_color=TERRA,
                                align="center", valign="vcenter", text_wrap=True))
        ws.write(13, i, en, k.f(font_name=EN, font_size=6.5, font_color="#F6D9CE",
                                bg_color=TERRA, align="center"))
    lab = k.f(font_name=FA, font_size=9, bg_color="#FFF1D6", border=1, border_color=LINE,
              align="right")
    cnt = k.f(font_name=EN, font_size=9, bg_color=WHITE, border=1, border_color=LINE_SOFT,
              align="center", num_format=INT)
    in_f = k.f(font_name=EN, font_size=9, bg_color=MINT, border=1, border_color=LINE_SOFT,
               align="center", num_format=INT)
    out_f = k.f(font_name=EN, font_size=9, bg_color=BLUSH, border=1, border_color=LINE_SOFT,
                align="center", num_format=INT)
    net_f = k.f(font_name=EN, font_size=9, bold=True, bg_color=CREAM, border=1,
                border_color=LINE_SOFT, align="center", num_format=INT_RED)
    bar = k.f(font_name=EN, font_size=7, font_color=TERRA, bg_color=WHITE, align="right")
    r0 = 14
    for i, (code, title, scope, group) in enumerate(pers):
        r = r0 + i
        ws.write_string(r, 0, title, lab)
        ws.write_formula(r, 1, f'=COUNTIF(Txn!$F$13:$F${12+NR},$A{r+1})', cnt)
        ws.write_formula(r, 2, f'=SUMIFS(Txn!$M$13:$M${12+NR},Txn!$F$13:$F${12+NR},$A{r+1},'
                               f'Txn!$T$13:$T${12+NR},"IN")', in_f)
        ws.write_formula(r, 3, f'=SUMIFS(Txn!$M$13:$M${12+NR},Txn!$F$13:$F${12+NR},$A{r+1},'
                               f'Txn!$T$13:$T${12+NR},"OUT")', out_f)
        ws.write_formula(r, 4, f'=$C{r+1}-$D{r+1}', net_f)
        ws.write_formula(r, 5, f'=REPT("█",MIN(14,ROUND($D{r+1}/'
                               f'MAX(1,MAX($D${r0+1}:$D${r0+n}))*14,0)))', bar)
    tot = r0 + n
    ws.write(tot, 0, "جمع", k.f(font_name=FA, font_size=10, bold=True,
                                font_color=WHITE, bg_color=TERRA, border=1))
    for c_ in (1, 2, 3, 4):
        ws.write_formula(tot, c_, f"=SUM({chr(65+c_)}{r0+1}:{chr(65+c_)}{tot})",
                         k.f(font_name=EN, font_size=9.5, bold=True, font_color=WHITE,
                             bg_color=TERRA, border=1, num_format=INT))

    # monthly
    m0 = tot + 3
    k.section(ws, m0, last, "۰۲", "روند ماهانه پول شخصی", "هر ماه چند تومان از شرکت خارج شد؟")
    mh = [("ماه", "Month", 9), ("برداشت شخصی", "Out", 16), ("برگشت", "In", 16),
          ("مانده", "Net", 16), ("تجمعی", "Cumulative", 16), ("نمودار", "Bar", 12)]
    for i, (fa, en, w) in enumerate(mh):
        ws.write(m0 + 1, i, fa, k.f(font_name=FA, font_size=9, bold=True, font_color=WHITE,
                                    bg_color=FOREST_2, border=1, border_color=FOREST_2,
                                    align="center"))
    for i in range(12):
        r = m0 + 2 + i
        ws.write_number(r, 0, i + 1, k.f(font_name=EN, font_size=9, bg_color=SAND_2, border=1,
                                         border_color=LINE, align="center", num_format="00"))
        ws.write_formula(r, 1, f'=SUMIFS(Txn!$M$13:$M${12+NR},Txn!$AD$13:$AD${12+NR},$A{r+1},'
                               f'Txn!$Z$13:$Z${12+NR},"شخصی رئیس",Txn!$T$13:$T${12+NR},"OUT")',
                         out_f)
        ws.write_formula(r, 2, f'=SUMIFS(Txn!$M$13:$M${12+NR},Txn!$AD$13:$AD${12+NR},$A{r+1},'
                               f'Txn!$Z$13:$Z${12+NR},"شخصی رئیس",Txn!$T$13:$T${12+NR},"IN")',
                         in_f)
        ws.write_formula(r, 3, f'=$B{r+1}-$C{r+1}', net_f)
        ws.write_formula(r, 4, f'=IF($A{r+1}="","",N($D{r+1}))' if i == 0 else
                         f'=IF($A{r+1}="","",N($E{r})+N($D{r+1}))', net_f)
        ws.write_formula(r, 5, f'=REPT("█",MIN(12,ROUND($B{r+1}/'
                               f'MAX(1,MAX($B${m0+2}:$B${m0+13}))*12,0)))', bar)
    ch = wb.add_chart({"type": "column"})
    ch.add_series({"name": "برداشت شخصی", "categories": ["Personal", m0 + 2, 0, m0 + 13, 0],
                   "values": ["Personal", m0 + 2, 1, m0 + 13, 1],
                   "fill": {"color": TERRA}, "gap": 55})
    ch.add_series({"name": "برگشت", "categories": ["Personal", m0 + 2, 0, m0 + 13, 0],
                   "values": ["Personal", m0 + 2, 2, m0 + 13, 2],
                   "fill": {"color": SAGE}, "gap": 55})
    ch.set_title({"name": "ماه به ماه", "name_font": {"name": FA, "size": 10}})
    ch.set_legend({"position": "bottom", "font": {"name": FA, "size": 8}})
    ch.set_size({"width": 360, "height": 230})
    ws.insert_chart(f"F{m0+2}", ch, {"x_offset": 6, "y_offset": 4})
    note_row = m0 + 16
    k.note(ws, note_row, last,
           "برای تسویه با رئیس: هر بار پول شخصی برداشت می‌شود «پرداخت به حساب شخصی رئیس» و هر بار "
           "برمی‌گرداند «دریافت از حساب شخصی رئیس» را در شیت Txn ثبت کنید. مانده این شیت همان "
           "بدهی رئیس به شرکت است و در فایل تراز (05_Bal) در حساب جاری رئیس دیده می‌شود.",
           bg=SAND, icon="people", h=48)
    ws.print_area(0, 0, note_row, last)
    return ws


# ===================================================================  Payable
def sheet_payable(wb, k, bd):
    """بدهی/طلب هر شخص «باسکول دیگران» و ردیابی آن از دفتر بانک."""
    ws = wb.add_worksheet("Payable")
    n = max(len(bd), 1)
    last = 6
    end = 12 + n + 12
    k.setup(ws, last, end, landscape=False, repeat=(11, 11), freeze=(13, 3))
    k.brand(ws, last, "بدهی و طلب «باسکول دیگران»",
            "مانده هر شخص از دفتر تراکنش‌های بانکی  ·  دریافت و پرداخت  ·  گزارش قابل چاپ",
            icon="people")
    k.kpis(ws, 6, [
        ("جمع بستانکار (بدهی ایجادشده)", "=SUM($D$13:$D$%d)" % (12 + n), INT, TERRA, "money"),
        ("جمع بدهکار (تسویه‌شده)", "=SUM($C$13:$C$%d)" % (12 + n), INT, GREEN, "money"),
        ("مانده بدهی ما", "=SUM($D$13:$D$%d)-SUM($C$13:$C$%d)" % (12 + n, 12 + n), INT, FOREST, "gauge"),
        ("اشخاص دارای مانده", '=COUNTIF($E$13:$E$%d,"<>0")' % (12 + n), INT, GOLD, "people"),
    ], last)
    k.section(ws, 9, 5, "۰۱", "بدهی و طلب به تفکیک شخص",
              "مانده = بستانکار − بدهکار  ·  ردیابی از دفتر بانک")
    heads = [("شخص / باسکول دیگران", "Person", 30), ("کد", "Code", 9),
             ("بدهکار (تسویه‌شده)", "Debit", 17), ("بستانکار (بدهی ایجادشده)", "Credit", 18),
             ("مانده بدهی ما", "Balance", 17), ("وضعیت", "Status", 20)]
    for i, (fa, en, w) in enumerate(heads):
        ws.set_column(i, i, w)
        ws.write(11, i, fa, k.f(font_name=FA, font_size=9.5, bold=True, font_color=WHITE,
                                bg_color=FOREST, border=1, border_color=FOREST,
                                align="center", valign="vcenter"))
    ws.set_row(11, 30)
    for i, person in enumerate(bd):
        r = 12 + i
        code = person[0]
        name = person[2] if len(person) > 2 else person[0]
        ws.set_row(r, 20)
        ws.write_formula(r, 0, f"=IFERROR(VLOOKUP($B{r+1},AcctTbl,2,0),$B{r+1})",
                         k.f(font_name=FA, font_size=9.5, bold=True, font_color=FOREST,
                             bg_color=SAND_2, border=1, border_color=LINE, align="right"),
                         name)
        ws.write_string(r, 1, code, k.f(font_name=EN, font_size=8.5, bg_color=SAND_2,
                                        border=1, border_color=LINE, align="center"))
        ws.write_formula(r, 2,
                         f'=SUMIF(Post!$D$13:$D${12+PN},$B{r+1},Post!$I$13:$I${12+PN})',
                         k.f(font_name=EN, font_size=9, bg_color=MINT, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(r, 3,
                         f'=SUMIF(Post!$D$13:$D${12+PN},$B{r+1},Post!$J$13:$J${12+PN})',
                         k.f(font_name=EN, font_size=9, bg_color=BLUSH, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(r, 4, f"=$D{r+1}-$C{r+1}",
                         k.f(font_name=EN, font_size=10, bold=True, bg_color=CREAM, border=1,
                             border_color=LINE_SOFT, align="center", num_format=INT_RED))
        ws.write_formula(r, 5,
                         f'=IF(ROUND($E{r+1},0)=0,"✓ تسویه‌شده",IF($E{r+1}>0,"بستانکار — بدهی ما","بدهکار — طلب ما"))',
                         k.f(font_name=FA, font_size=8.5, bg_color=WHITE, border=1,
                             border_color=LINE_SOFT, align="center"))
    tot = 12 + n
    ws.merge_range(tot, 0, tot, 1, "جمع", k.f(font_name=FA, font_size=10, bold=True,
                                              font_color=WHITE, bg_color=GOLD, border=1))
    for c in (2, 3, 4):
        ws.write_formula(tot, c, f"=SUM({chr(65+c)}13:{chr(65+c)}{tot})",
                         k.f(font_name=EN, font_size=10, bold=True, font_color=WHITE,
                             bg_color=GOLD, border=1, num_format=INT_RED))
    ws.write_formula(tot, 5, f'=IF(ROUND($E{tot+1},0)=0,"✓ جمع صفر","جمع بدهی ما")',
                     k.f(font_name=FA, font_size=9.5, bold=True, font_color=WHITE,
                         bg_color=GOLD, border=1))

    ch = wb.add_chart({"type": "bar"})
    ch.add_series({"name": "مانده", "categories": ["Payable", 12, 0, 11 + n, 0],
                   "values": ["Payable", 12, 4, 11 + n, 4],
                   "fill": {"color": TERRA}, "gap": 40,
                   "data_labels": {"value": True, "num_format": "#,##0",
                                   "font": {"size": 7}}})
    ch.set_title({"name": "مانده هر شخص باسکول دیگران",
                  "name_font": {"name": FA, "size": 11}})
    ch.set_legend({"none": True})
    ch.set_plotarea({"fill": {"color": "#FCFAF4"}})
    ch.set_size({"width": 460, "height": max(240, 26 * n)})
    ws.insert_chart("H12", ch, {"x_offset": 8, "y_offset": 6})

    k.note(ws, tot + 2, 6,
           "هر «دریافت هزینه باسکول دیگران» این شخص را بستانکار (بدهی شرکت) و هر «پرداخت به باسکول دیگران» "
           "او را بدهکار می‌کند. برای افزودن شخص جدید: 00_Master ← شیت Acc ← یک کد BD خالی را پر کنید؛ "
           "همان لحظه در کشوی فایل بانک می‌آید و همین گزارش آن را نشان می‌دهد.", bg=SAND, icon="people")
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
            f'=SUMIFS(Txn!$O$13:$O${12+NR},Txn!$J$13:$J${12+NR},INDEX(BankTbl,{i+1},1),'
            f'Txn!$L$13:$L${12+NR},"کارت/پوز")',
            k.f(font_name=EN, font_size=9, bg_color=MINT, border=1,
                border_color=LINE_SOFT, align="center", num_format=INT))
        ws.write_formula(
            r, 3,
            f'=SUMIFS(Txn!$O$13:$O${12+NR},Txn!$J$13:$J${12+NR},INDEX(BankTbl,{i+1},1),'
            f'Txn!$E$13:$E${12+NR},"واریز پز باسکول (۲۰۰۶)")',
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
    k.setup(ws, last, 12 + NCHQ + 3, landscape=True, repeat=(11, 11), freeze=(12, 2))
    k.brand(ws, last, "دفتر چک‌ها",
            "چک‌های دریافتی و پرداختی  ·  سررسید، وضعیت و مبلغ در یک نگاه", icon="cheque")
    k.kpis(ws, 6, [
        ("تعداد چک", f"=COUNTA($C$13:$C${12+NCHQ})", INT, FOREST, "cheque"),
        ("در جریان", f'=SUMIFS($I$13:$I${12+NCHQ},$J$13:$J${12+NCHQ},"در جریان")', INT, GOLD, "calendar"),
        ("پاس شده", f'=SUMIFS($I$13:$I${12+NCHQ},$J$13:$J${12+NCHQ},"پاس شده")', INT, GREEN, "check"),
        ("برگشتی", f'=SUMIFS($I$13:$I${12+NCHQ},$J$13:$J${12+NCHQ},"برگشتی")', INT, RED, "warn"),
        ("جمع کل", f"=SUM($I$13:$I${12+NCHQ})", INT, BLUE, "money"),
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
    for i in range(NCHQ):
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
    ws.data_validation(12, 4, 12 + NCHQ - 1, 4, {"validate": "list", "source": "=ListChequeTypes",
                                     "error_type": "stop"})
    ws.data_validation(12, 7, 12 + NCHQ - 1, 7, {"validate": "list", "source": "=ListBanks",
                                      "error_type": "warning"})
    ws.data_validation(12, 9, 12 + NCHQ - 1, 9, {"validate": "list", "source": "=ListChequeStatus"})
    ws.data_validation(12, 5, 12 + NCHQ - 1, 5, {"validate": "list", "source": "=ListAllCodes",
                                      "error_type": "warning"})
    ws.conditional_format(12, 9, 12 + NCHQ - 1, 9, {
        "type": "text", "criteria": "containing", "value": "پاس",
        "format": k.f(bg_color=MINT, font_color=GREEN, bold=True)})
    ws.conditional_format(12, 9, 12 + NCHQ - 1, 9, {
        "type": "text", "criteria": "containing", "value": "برگشتی",
        "format": k.f(bg_color="#FBE3D8", font_color=TERRA, bold=True)})
    ws.conditional_format(12, 9, 12 + NCHQ - 1, 9, {
        "type": "text", "criteria": "containing", "value": "در جریان",
        "format": k.f(bg_color=SAND, font_color=GOLD, bold=True)})
    ws.autofilter(11, 0, 12 + NCHQ - 1, 11)
    return ws


# =====================================================================  Home
def sheet_home(wb, k, acc, banks):
    ws = wb.add_worksheet("Home")
    last = 8
    k.setup(ws, last, 40, landscape=True, repeat=None, freeze=None)
    k.brand(ws, last, "فایل تراکنش‌های بانکی",
            "ثبت، کنترل و تحلیل گردش حساب‌های شرکت  ·  همه کشوها از فایل مادر تعریف شده‌اند", icon="bank")
    k.kpis(ws, 6, [
        ("تراکنش‌ها", f"=COUNTA(Txn!$H$13:$H${12+NR})", INT, FOREST, "list"),
        ("دریافت", f"=SUM(Txn!$O$13:$O${12+NR})", INT, GREEN, "transfer"),
        ("پرداخت", f"=SUM(Txn!$P$13:$P${12+NR})", INT, RED, "transfer"),
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
        ("طبقه‌بندی  Class", "پول برای چه کاری رفت؟ گردش هر طبقه‌بندی، شرکتی و شخصی",
         "internal:Class!A1", "list", "#4F7C6C"),
        ("شخصی رئیس  Personal", "مصرف شخصی رئیس از حساب شرکت و بدهی او به شرکت",
         "internal:Personal!A1", "people", TERRA),
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
                       k.f(font_name=FA, font_size=10, font_color=INK, align="right",
                           valign="vcenter", border=1, border_color=LINE_SOFT))
        ws.set_row(r, 24)
        r += 1
    k.note(ws, r + 1, last,
           "قاعده طلایی: هر بانکی که در فهرست بانک‌های شرکت نیست، اجازه ثبت ندارد. اگر حساب جدیدی باز کردید، "
           "اول آن را در فایل مادر (شیت Bnk) اضافه کنید و بعد این فایل را باز کنید.", bg=MINT, icon="flag")
    k.section(ws, r + 3, last, "۰۳", "فایل‌های همراه",
              "همه در همین پوشه‌اند؛ با مسیر نسبی باز می‌شوند")
    k.nav(ws, r + 4, [
        ("فایل مادر  00_Master", "تعریف حساب‌ها، بانک‌ها، اشخاص باسکول دیگران و همه کشوها",
         "external:00_Master.xlsx", "list", FOREST),
        ("دفتر ثبت  01_Entry", "ثبت روزانه باسکول، سبد و پرداخت‌ها؛ منبع دفتر روزنامه",
         "external:01_Entry.xlsx", "list", GOLD),
        ("داشبورد  08_Dash", "نمای کلی بانک‌ها، پز و محصولات با نمودار زنده",
         "external:08_Dash.xlsx", "chart", SAGE),
        ("تراز  05_Bal", "تراز آزمایشی و دفتر حساب‌ها",
         "external:05_Bal.xlsx", "scale", BLUE),
        ("گزارش‌ها  03_Report", "گزارش سود و زیان و گردش حساب‌ها",
         "external:03_Report.xlsx", "gauge", TERRA),
        ("راهنمای این فایل  Help", "آموزش گام‌به‌گام شیت‌ها و قواعد ثبت",
         "internal:Help!A1", "help", "#8E6C88"),
    ], last, per_row=3)
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
    classes = read_classes(master_path)
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
    sheet_class(wb, k, classes)
    sheet_personal(wb, k, classes)
    sheet_payable(wb, k, bd)
    sheet_card(wb, k, banks)
    sheet_chq(wb, k)
    sheet_sync(wb, k, acc, banks, gl, bd, classes)
    sheet_help(wb, k, acc, banks)
    wb.worksheets()[0].activate()
    wb.close()
    return out_path
