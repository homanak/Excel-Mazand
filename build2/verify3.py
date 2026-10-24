# -*- coding: utf-8 -*-
"""Health-check for the Mazandasht v03 folder.

    /home/user/.venv/bin/python build2/verify3.py [--deep]

Checks, per workbook: every sheet exists in the package map, is right-to-left,
A4 portrait/landscape with fit-to-width and a header/footer.
With --deep it also recalculates 02_Bank.xlsx (formulas) and prints the key totals.
"""
import os
import re
import sys
import warnings
import zipfile

import openpyxl

warnings.filterwarnings("ignore")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "Mazandasht_v03")
FILES = ["00_Master.xlsx", "01_Entry.xlsx", "02_Bank.xlsx", "03_Report.xlsx",
         "04_Inv.xlsx", "05_Bal.xlsx", "06_Buyers.xlsx", "07_Demo.xlsx",
         "08_Dash.xlsx"]
EXPECT = {
    "00_Master.xlsx": ["Home", "List", "Sync", "Help"],
    "02_Bank.xlsx": ["Home", "Txn", "Post", "Sum", "Flow", "Anal", "Card", "Chq", "Sync", "Help"],
    "08_Dash.xlsx": ["Dash", "Pos", "Banks", "Time", "Ops", "Parties", "F1", "F2", "Sync", "Help"],
}

FMT = "{:<18s} {:>3d} sheets  {:>3d} issues  {}"


def flags(z, name):
    wb_xml = z.read("xl/workbook.xml").decode()
    rels = z.read("xl/_rels/workbook.xml.rels").decode()
    rmap = {}
    for m in re.finditer(r"<Relationship\b([^>]*)/>", rels):
        attrs = dict(re.findall(r'(\w+)="([^"]*)"', m.group(1)))
        if "Id" in attrs and "Target" in attrs:
            rmap[attrs["Id"]] = attrs["Target"]
    rid = dict(re.findall(r'<sheet name="([^"]+)"[^>]*?r:id="(rId\d+)"', wb_xml)).get(name)
    if not rid or rid not in rmap:
        return None
    target = rmap[rid]
    path = target.lstrip("/") if target.startswith("/") else "xl/" + target
    xml = z.read(path).decode()
    return {
        "rtl": 'rightToLeft="1"' in xml,
        "paper": (re.search(r'paperSize="(\d+)"', xml) or [None, "-"])[1],
        "fit": "fitToPage" in xml,
        "hf": "<headerFooter" in xml,
    }


def check_book(fname, deep=False):
    path = os.path.join(OUT, fname)
    if not os.path.exists(path):
        return None, [f"{fname} missing"]
    issues = []
    with zipfile.ZipFile(path) as z:
        wb = openpyxl.load_workbook(path)
        names = wb.sheetnames
        for sheet in names:
            f = flags(z, sheet)
            if f is None:
                issues.append(f"{sheet}:not-in-package")
                continue
            if not f["rtl"]:
                issues.append(f"{sheet}:no-rtl")
            if f["paper"] != "9":
                issues.append(f"{sheet}:paper={f['paper']}")
            if not f["fit"]:
                issues.append(f"{sheet}:no-fit")
            if not f["hf"]:
                issues.append(f"{sheet}:no-header")
        for need in EXPECT.get(fname, []):
            if need not in names:
                issues.append(f"missing-sheet {need}")
        if fname == "08_Dash.xlsx":
            ext = [n for n in z.namelist() if n.startswith("xl/externalLinks/externalLink")]
            if not any(n.endswith(".xml") and "_rels" not in n for n in ext):
                issues.append("no-external-links")
            else:
                body = z.read("xl/externalLinks/externalLink1.xml").decode()
                if "<sheetData" not in body:
                    issues.append("empty-link-cache")
        if fname == "02_Bank.xlsx":
            ws = wb["Txn"]
            dvs = {str(dv.sqref): dv.formula1 for dv in ws.data_validations.dataValidation}
            if not any(v == "ListBanks" for v in dvs.values()):
                issues.append("Txn:no-bank-list")
            if not any(v == "ListAllCodes" for v in dvs.values()):
                issues.append("Txn:no-party-code")
            if not any(v == "ListOtherScales" for v in dvs.values()):
                issues.append("Txn:no-other-scale")
        wb.close()
    return len(names), issues


def deep_numbers():
    import formulas
    path = os.path.join(OUT, "02_Bank.xlsx")
    model = formulas.ExcelModel().loads(path).finish()
    sol = model.calculate()
    get = {}
    for k, v in sol.items():
        val = getattr(v, "value", v)
        if hasattr(val, "item") and not hasattr(val, "__len__"):
            try:
                val = val.item()
            except Exception:
                val = str(val)
        get[k] = val
    key = lambda s, r: get.get(f"'[02_Bank.xlsx]{s}'!{r}")
    errs = [k for k, v in get.items() if isinstance(v, str) and v.startswith("#")]
    rows = {"total_balance": key("SUM", "E22"), "inflow": key("SUM", "C22"),
            "outflow": key("SUM", "D22"), "flow_net": key("FLOW", "E25"),
            "card_2006": key("CARD", "C14"), "pos_2006": key("CARD", "D14"),
            "anal_in": key("ANAL", "C10"), "txn_status_ok": key("TXN", "S13")}
    print("\n02_Bank.xlsx recalculation")
    print("  errors:", len(errs), errs[:3])
    for name, val in rows.items():
        print(f"  {name:16s} {val}")
    return rows


def main():
    print("Mazandasht v03 — health check\n")
    total = 0
    for fname in FILES:
        n, issues = check_book(fname)
        if n is None:
            print(FMT.format(fname, 0, len(issues), "; ".join(issues)))
            total += len(issues)
            continue
        print(FMT.format(fname, n, len(issues), "; ".join(issues[:4])))
        total += len(issues)
    print(f"\nissues: {total}")
    if "--deep" in sys.argv:
        deep_numbers()
    return 0 if total == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
