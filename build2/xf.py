# -*- coding: utf-8 -*-
"""XML-level transformations for the Mazandasht v02 series.

Everything here works on the OOXML package directly (zip + text), so the
external-link parts, comments, validations and styles that xlsxwriter made
stay exactly as they are; we only rename sheets, force RTL, and add print
setup.
"""
import os
import re
import shutil
import zipfile

# ----------------------------------------------------------------- file map
FILES = {
    "INFO.xlsx": "00_Master.xlsx",
    "01_Register.xlsx": "01_Entry.xlsx",
    "02_Bank.xlsx": "02_Bank.xlsx",
    "03_Report.xlsx": "03_Report.xlsx",
    "04_Invoice.xlsx": "04_Inv.xlsx",
    "05_Balance.xlsx": "05_Bal.xlsx",
    "06_LoadBuyers.xlsx": "06_Buyers.xlsx",
    "07_Sample.xlsx": "07_Demo.xlsx",
}

# Sheets used by the reports from the entry file (renamed in every file).
REG_SHEETS = {
    "Baskol": "Scale",
    "Basket": "Crate",
    "Payment": "Pay",
    "Ledger": "Book",
    "Summary": "Sum",
}

SHEETS = {
    "00_Master.xlsx": {
        "Company": "Co", "Params": "Par", "Groups": "Grp", "Accounts": "Acc",
        "Ops": "Op", "Baskets": "Bsk", "Products": "Prd", "Banks": "Bnk",
        "Lists": "List",
    },
    "01_Entry.xlsx": dict(REG_SHEETS),
    "02_Bank.xlsx": {
        "Overview": "Ov", "Transactions": "Txn", "Cheques": "Chq",
        "Reconcile": "Rec", "Summary": "Sum",
        "Baskol": "Scale", "Basket": "Crate", "Payment": "Pay",
        "Ledger": "Book",
    },
    "03_Report.xlsx": {
        "Overview": "Ov", "Growers": "Grw", "LoadBuyers": "LBuy",
        "BasketSellers": "BSel", "Groups": "Grp",
        "Baskol": "Scale", "Basket": "Crate", "Payment": "Pay",
        "Ledger": "Book", "Summary": "Sum",
    },
    "04_Inv.xlsx": {
        "Invoice": "Inv",
        "Baskol": "Scale", "Basket": "Crate", "Payment": "Pay",
        "Ledger": "Book", "Summary": "Sum",
    },
    "05_Bal.xlsx": {
        "Overview": "Ov", "Balance4": "B4", "Balance6": "B6",
        "Growers": "Grw", "LoadBuyers": "LBuy", "BasketSellers": "BSel",
        "Inventory": "Stock", "Interbank": "IBank",
        "Baskol": "Scale", "Basket": "Crate", "Payment": "Pay",
        "Ledger": "Book", "Summary": "Sum",
    },
    "06_Buyers.xlsx": {
        "Overview": "Ov", "Detail": "Det", "Monthly": "Mon",
        "Baskol": "Scale", "Basket": "Crate", "Payment": "Pay",
        "Ledger": "Book", "Summary": "Sum",
    },
    "07_Demo.xlsx": {
        "Register": "Reg", "Summary": "Sum",
    },
}

TITLES = {
    "00_Master.xlsx": "MASTER DATA",
    "01_Entry.xlsx": "DAILY ENTRY",
    "02_Bank.xlsx": "BANK & CHEQUES",
    "03_Report.xlsx": "REPORTS",
    "04_Inv.xlsx": "INVOICE",
    "05_Bal.xlsx": "BALANCE SHEET",
    "06_Buyers.xlsx": "LOAD BUYERS",
    "07_Demo.xlsx": "DEMO",
}

# Strings that must follow the new file names inside cells / help texts.
TEXT_SWAPS = [
    ("01_Register.xlsx", "01_Entry.xlsx"),
    ("INFO.xlsx", "00_Master.xlsx"),
    ("04_Invoice.xlsx", "04_Inv.xlsx"),
    ("05_Balance.xlsx", "05_Bal.xlsx"),
    ("06_LoadBuyers.xlsx", "06_Buyers.xlsx"),
    ("07_Sample.xlsx", "07_Demo.xlsx"),
]


def _ref_re(old):
    """Matches SheetName in `SheetName!` and in `'SheetName'!` only."""
    return re.compile(r"(?<![A-Za-z0-9_.])" + re.escape(old) + r"(?=!|'!)")


def rename_sheets(text, mapping):
    for old, new in sorted(mapping.items(), key=lambda kv: -len(kv[0])):
        if old == new:
            continue
        text = _ref_re(old).sub(new, text)
        text = text.replace(f'name="{old}"', f'name="{new}"')
        text = text.replace(f'val="{old}"', f'val="{new}"')
    return text


def swap_text(text):
    for old, new in TEXT_SWAPS:
        text = text.replace(old, new)
    return text


# ------------------------------------------------------------ broken formulas
def fix_formulas(xml):
    """Repairs three defects of the v01 report files:
       $E11$      -> $E$11      (invalid absolute row)
       $C10=,     -> $C10="",   (missing empty string)
       ),,N(      -> ),"",N(    (missing empty string)
    """
    before = xml
    xml = re.sub(r"\$([A-Z]{1,3})(\d+)\$", r"$\1$\2", xml)
    xml = re.sub(r"(\$[A-Z]{1,3}\d+)=,", r'\1="",', xml)
    xml = xml.replace("),,N(", '),"",N(')
    return xml, xml != before


# --------------------------------------------------------------- print setup
HEADER = {
    "00_Master.xlsx": "MASTER DATA",
    "01_Entry.xlsx": "DAILY ENTRY",
    "02_Bank.xlsx": "BANK",
    "03_Report.xlsx": "REPORTS",
    "04_Inv.xlsx": "INVOICE",
    "05_Bal.xlsx": "BALANCE",
    "06_Buyers.xlsx": "BUYERS",
    "07_Demo.xlsx": "DEMO",
}


def _header_footer(fname):
    title = HEADER[fname]
    head = ("&amp;Lمازندشت — سامانه حسابداری سورتینگ"
            f"&amp;C&amp;\"B Nazanin,Bold\"&amp;10 MAZANDASHT · {title}"
            "&amp;R&amp;A")
    foot = ("&amp;L&amp;F"
            "&amp;Cصفحه &amp;P از &amp;N"
            "&amp;RMAZANDASHT · Excel 2013")
    return (f"<headerFooter><oddHeader>{head}</oddHeader>"
            f"<oddFooter>{foot}</oddFooter></headerFooter>")


def print_setup(xml, fname):
    # ---- sheetPr + fitToPage -------------------------------------------
    if "<pageSetUpPr" not in xml:
        m = re.search(r"<sheetPr\b([^>]*?)/>", xml)
        if m:
            xml = (xml[:m.start()] + f"<sheetPr{m.group(1)}>"
                   '<pageSetUpPr fitToPage="1"/></sheetPr>' + xml[m.end():])
        else:
            m = re.search(r"<sheetPr\b[^>]*>", xml)
            if m:
                xml = xml[:m.end()] + '<pageSetUpPr fitToPage="1"/>' + xml[m.end():]
            else:
                m = re.search(r"<worksheet\b[^>]*>", xml)
                xml = (xml[:m.end()] + '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>'
                       + xml[m.end():])

    # ---- right to left --------------------------------------------------
    m = re.search(r"<sheetView\b([^>]*?)(/?)>", xml)
    if m:
        attrs, slash = m.group(1), m.group(2)
        if "rightToLeft" not in attrs:
            xml = (xml[:m.start()] + f"<sheetView{attrs} rightToLeft=\"1\"{slash}>"
                   + xml[m.end():])
    else:
        xml = re.sub(r"<sheetViews\s*/>",
                     '<sheetViews><sheetView rightToLeft="1" workbookViewId="0"/></sheetViews>',
                     xml, count=1)

    # ---- A4 + fit to width ---------------------------------------------
    m = re.search(r"<pageSetup\b([^>]*?)/>", xml)
    if m:
        attrs = m.group(1)
        attrs = re.sub(r'\s*paperSize="\d+"', "", attrs)
        attrs = re.sub(r'\s*fitToWidth="\d+"', "", attrs)
        attrs = re.sub(r'\s*fitToHeight="\d+"', "", attrs)
        if "orientation" not in attrs:
            attrs += ' orientation="landscape"'
        xml = (xml[:m.start()]
               + f'<pageSetup{attrs} paperSize="9" fitToWidth="1" fitToHeight="0"/>'
               + xml[m.end():])
    else:
        block = '<pageSetup paperSize="9" fitToWidth="1" fitToHeight="0" orientation="landscape"/>'
        xml = _insert_before_close(xml, block)

    # ---- header / footer ------------------------------------------------
    if "<headerFooter" not in xml:
        xml = _insert_before_close(xml, _header_footer(fname))
    elif "<oddHeader" not in xml or "<oddHeader />" in xml or "<oddHeader/>" in xml:
        hf = _header_footer(fname)
        head = re.search(r"<oddHeader>(.*?)</oddHeader>", hf).group(1)
        xml = xml.replace("<oddHeader />", "").replace("<oddHeader/>", "")
        m = re.search(r"<headerFooter[^>]*>", xml)
        xml = xml[:m.end()] + f"<oddHeader>{head}</oddHeader>" + xml[m.end():]
    return xml


def _insert_before_close(xml, block):
    idx = [xml.find(a) for a in ("<legacyDrawing", "<drawing") if xml.find(a) > 0]
    i = min(idx) if idx else xml.find("</worksheet>")
    return xml[:i] + block + xml[i:]


# ------------------------------------------------------------------ transform
def transform_package(src, dst, fname):
    sheet_map = SHEETS[fname]
    tmp = dst + ".tmp"
    with zipfile.ZipFile(src) as zin:
        items = {n: zin.read(n) for n in zin.namelist()}

    fixed = 0
    for name, data in list(items.items()):
        if not (name.endswith(".xml") or name.endswith(".rels")):
            continue
        try:
            xml = data.decode("utf-8")
        except UnicodeDecodeError:
            continue
        orig = xml

        is_sheet = re.fullmatch(r"xl/worksheets/sheet\d+\.xml", name) is not None
        if is_sheet or name in ("xl/workbook.xml", "docProps/app.xml") or name.startswith("xl/externalLinks/"):
            xml = rename_sheets(xml, sheet_map)
        if is_sheet:
            xml, did = fix_formulas(xml)
            fixed += 1 if did else 0
            xml = print_setup(xml, fname)
        if name in ("xl/sharedStrings.xml", "xl/workbook.xml", "docProps/app.xml") or \
                name.startswith("xl/worksheets/"):
            xml = swap_text(xml)
        xml = xml.replace("01_Register.xlsx", "01_Entry.xlsx")
        if xml != orig:
            items[name] = xml.encode("utf-8")

    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in items.items():
            zout.writestr(name, data)
    shutil.move(tmp, dst)
    return fixed


def run(src_dir, dst_dir):
    os.makedirs(dst_dir, exist_ok=True)
    out = []
    for old, new in FILES.items():
        src = os.path.join(src_dir, old)
        dst = os.path.join(dst_dir, new)
        if not os.path.exists(src):
            print("  missing", src)
            continue
        fixed = transform_package(src, dst, new)
        out.append((new, os.path.getsize(dst), fixed))
        print(f"  {new:<18} {os.path.getsize(dst)//1024:>4} KB   formulas fixed: {fixed}")
    return out


if __name__ == "__main__":
    here = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    run(os.path.join(here, "Mazandasht_System_v01"),
        os.path.join(here, "Mazandasht_v02"))
