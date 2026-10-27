# -*- coding: utf-8 -*-
"""«INFO» دیگر وجود ندارد — نام و متن‌های قدیمی را در همهٔ فایل‌ها یکسان می‌کند.

The first versions called the mother file ``INFO.xlsx`` and the entry file
``01_Register``; v04 renamed them to ``00_Master.xlsx`` / ``01_Entry.xlsx`` but
the tutorial text inside the workbooks was inherited from the old seeds, so
``Home`` and ``Help`` still told the user to «اول INFO را باز کنید» and a few
home-page cards were labelled with the old file names.

``fix_book(wb)`` rewrites those strings and ``fix_all(folder)`` applies it to
the nine delivered workbooks.
"""
import os
import re

import openpyxl

ACC = "00_Master"

# order matters: the long forms first
RULES = [
    (r"فایل مادر\s*INFO", f"فایل مادر {ACC}"),
    (r"INFO\s*/\s*Ops", f"{ACC}/Op"),
    (r"INFO\s*/\s*Baskets", f"{ACC}/Bsk"),
    (r"INFO\s*/\s*Products", f"{ACC}/Prd"),
    (r"INFO\s*/\s*Customers", f"{ACC}/Acc"),
    (r"INFO\.xlsx", f"{ACC}.xlsx"),
    (r"\bINFO\b", f"فایل مادر {ACC}"),
    (r"۰۱_Register|01_Register", "01_Entry"),
    (r"04_Invoice", "04_Inv"),
    (r"\bRegister\b", "ثبت"),
]

# home-page cards that pointed at the wrong thing after the rename
LABELS = {
    "باز کردن فایل مادر INFO": f"فایل مادر — {ACC}.xlsx  (حساب‌ها، بانک‌ها، فهرست کشوها)",
    f"باز کردن فایل مادر {ACC}": f"فایل مادر — {ACC}.xlsx  (حساب‌ها، بانک‌ها، فهرست کشوها)",
    "بارفروشان  ·  06_Buyers.xlsx": "تحلیل خریداران  ·  06_Buyers.xlsx",
    "ریز بار، ماهانه و مانده": "تحلیل خریداران، ماهانه، مانده و وضعیت",
}


def _label(text):
    """Exact card labels (whitespace-insensitive)."""
    key = re.sub(r"\s+", " ", text).strip()
    for k, v in LABELS.items():
        if re.sub(r"\s+", " ", k).strip() == key:
            return v
    return None


def _sub(text):
    lab = _label(text)
    if lab:
        return lab
    out = text
    for pat, rep in RULES:
        out = re.sub(pat, rep, out)
    return _label(out) or out


def fix_book(wb):
    n = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and not v.startswith("=") and (
                        "INFO" in v
                        or re.search(r"۰۱_Register|01_Register|04_Invoice|\bRegister\b", v)):
                    new = _sub(v)
                    if new != v:
                        c.value = new
                        n += 1
                elif isinstance(v, str) and not v.startswith("=") and _label(v):
                    c.value = _label(v)
                    n += 1
    return n


def fix_all(folder):
    total = {}
    for name in sorted(os.listdir(folder)):
        if not name.endswith(".xlsx"):
            continue
        path = os.path.join(folder, name)
        wb = openpyxl.load_workbook(path)
        n = fix_book(wb)
        if n:
            wb.calculation.fullCalcOnLoad = True
            wb.save(path)
        wb.close()
        total[name] = n
    return total


if __name__ == "__main__":
    import sys
    folder = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "Mazandasht_v04")
    print(fix_all(folder))
