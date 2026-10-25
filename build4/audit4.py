# -*- coding: utf-8 -*-
"""Structural audit of the v04 deliverable — reads the .xlsx XML, no recalc.

    python build4/audit4.py [--verbose]

The 50k-row books cannot be recalculated by `formulas`, so this audit verifies
the *structure*: A4 print logic on every sheet, RTL, headers/footers, print
areas, data validations (dropdowns), capacities, mirror sizes, cross-file
links, the bank sub-ledger and the payable (باسکول دیگران) report.
"""
import os
import re
import sys
import zipfile
from collections import Counter, defaultdict
from xml.etree import ElementTree as ET

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
OUT = os.path.join(ROOT, "Mazandasht_v04")

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}

OK, BAD = "✓", "✗"
issues = []


def bad(msg):
    issues.append(msg)


def sheets_of(zf):
    """[(name, xml path inside the zip)] in workbook order."""
    wb = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel = {e.get("Id"): e.get("Target") for e in rels}
    out = []
    for sh in wb.find("m:sheets", NS):
        rid = sh.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        tgt = rel[rid].lstrip("/")
        if not tgt.startswith("xl/"):
            tgt = "xl/" + tgt
        out.append((sh.get("name"), tgt))
    return out


def defined_names(zf):
    """{(localSheetId or -1, name): ref}"""
    wb = ET.fromstring(zf.read("xl/workbook.xml"))
    dn = wb.find("m:definedNames", NS)
    names = {}
    if dn is not None:
        for e in dn:
            key = (int(e.get("localSheetId")) if e.get("localSheetId") is not None else -1,
                   e.get("name") or "")
            names[key] = e.text or ""
    return names


def audit_file(path, verbose=False):
    fname = os.path.basename(path)
    with zipfile.ZipFile(path) as zf:
        names = sheets_of(zf)
        dns = defined_names(zf)
        ext_links = len([n for n in zf.namelist() if n.startswith("xl/externalLinks/externalLink")])
        print(f"\n{fname}  —  {len(names)} sheet(s)")
        for name, target in names:
            try:
                xml = zf.read(target)
            except KeyError:
                bad(f"{fname}/{name}: missing xml")
                continue
            root = ET.fromstring(xml)
            tag = f"{fname}!{name}"
            ps = root.find("m:pageSetup", NS)
            po = root.find("m:printOptions", NS)
            hf = root.find("m:headerFooter", NS)
            sv = root.find("m:sheetViews/m:sheetView", NS)
            pr = root.find("m:sheetPr/m:pageSetUpPr", NS)
            state = (sv.get("state") if sv is not None else "") or ""
            if state == "hidden":
                if verbose:
                    print(f"   · {name:<10} hidden (skipped)")
                continue
            problems = []
            if ps is None:
                problems.append("no pageSetup")
            else:
                if ps.get("paperSize") != "9":
                    problems.append(f"paper {ps.get('paperSize')}")
                if ps.get("orientation") not in ("portrait", "landscape"):
                    problems.append("no orientation")
            if pr is None or pr.get("fitToPage") != "1":
                problems.append("fitToPage off")
            if po is None or po.get("horizontalCentered") != "1":
                problems.append("not centered")
            hdr = "".join((e.text or "") for e in hf.findall("m:oddHeader", NS)) if hf is not None else ""
            ftr = "".join((e.text or "") for e in hf.findall("m:oddFooter", NS)) if hf is not None else ""
            if hf is not None and not hdr and hf.get("oddHeader"):
                hdr = hf.get("oddHeader")
            if not hdr.strip():
                problems.append("no header")
            if not ftr.strip():
                problems.append("no footer")
            if sv is None or sv.get("rightToLeft") != "1":
                problems.append("not RTL")
            idx = [i for i, (n_, _) in enumerate(names) if n_ == name][0]
            area = dns.get((idx, "_xlnm.Print_Area")) or dns.get((-1, "_xlnm.Print_Area"))
            if not area:
                problems.append("no print area")
            dv = len(root.findall("m:dataValidations/m:dataValidation", NS))
            hl = len(root.findall("m:hyperlinks/m:hyperlink", NS))
            if problems:
                bad(f"{tag}: {', '.join(problems)}")
            if verbose:
                print(f"   · {name:<10} A4 {ps.get('orientation') if ps is not None else '?':<9}"
                      f" dv={dv:<3} links={hl:<3} area={area or '-'}")
        dv_total = sum(len(ET.fromstring(zf.read(t)).findall(
            "m:dataValidations/m:dataValidation", NS)) for _, t in names)
        print(f"   dropdowns (data validations) in file: {dv_total:<4} "
              f"live external links: {ext_links}")
    return dv_total


def book_structure():
    print("\n01_Entry.xlsx — capacities and ranges")
    wb = openpyxl.load_workbook(os.path.join(OUT, "01_Entry.xlsx"),
                                read_only=True, data_only=False)
    bk = wb["Book"]
    print(f"   Book rows: {bk.max_row}   cols: {bk.max_column}")
    if bk.max_row != 50010:
        bad(f"Book last row {bk.max_row} ≠ 50010")
    # sampling the last row of every block
    want = {10: "Scale", 40009: "Scale", 40010: "Pay", 45009: "Pay",
            45010: "Crate", 50009: "Crate"}
    seen = {}
    for r, row in enumerate(bk.iter_rows(min_row=1, max_row=50010, values_only=True), 1):
        if r in want:
            seen[r] = [str(v or "") for v in row[:16]]
    for r, src in want.items():
        blob = " ".join(seen.get(r, []))
        if f"{src}!" not in blob:
            bad(f"Book row {r} does not read from {src}")
    for sheet, last in (("Scale", 10009), ("Pay", 2509), ("Crate", 2509),
                        ("Inv", 2009)):
        ws = wb[sheet]
        if sheet == "Inv" and ws.max_row < 2009:
            bad(f"Inv has only {ws.max_row} rows")
        if sheet != "Inv" and ws.max_row != last:
            # the last row may be a total row
            print(f"   {sheet}: max_row {ws.max_row} (expected {last})")
            if ws.max_row < last:
                bad(f"{sheet} has only {ws.max_row} rows")
    # counts of registered rows
    cnt = {}
    for sheet, key_col, first in (("Scale", 2, 10), ("Pay", 2, 10), ("Crate", 2, 10)):
        ws = wb[sheet]
        n = 0
        for row in ws.iter_rows(min_row=first, max_row=(10009 if sheet == "Scale" else 2509),
                                min_col=key_col, max_col=key_col, values_only=True):
            if row[0] not in (None, ""):
                n += 1
        cnt[sheet] = n
    print("   registered rows:", ", ".join(f"{k} {v}" for k, v in cnt.items()))
    wb.close()

    wb = openpyxl.load_workbook(os.path.join(OUT, "01_Entry.xlsx"), data_only=False)
    for sheet, want in (("Scale", "T10009"), ("Crate", "M2509"), ("Pay", "R2509"),
                        ("Book", "Q50010")):
        area = str(wb[sheet].print_area or "").replace("$", "")
        if want not in area:
            bad(f"{sheet} print area {area} does not reach {want}")
    for sheet in ("Inv", "Print", "GRep"):
        ws = wb[sheet]
        print(f"   {sheet}: dims {ws.dimensions}  print {ws.print_area}  "
              f"dv {len(ws.data_validations.dataValidation)}")
    if "Bill" not in wb.sheetnames:
        bad("01_Entry: no Bill sheet")
    else:
        bl = wb["Bill"]
        area = str(bl.print_area or "").replace("$", "").replace("'", "")
        net = str(bl["H11"].value or "")
        print(f"   Bill: dims {bl.dimensions}  print {area}  dv {len(bl.data_validations.dataValidation)}")
        if not area.startswith("Bill!A1:K"):
            bad(f"Bill print area {area} is not A1:K…")
        if "FLOOR(" not in net or "$O$5" not in net:
            bad("Bill net-kg formula missing FLOOR / loss cell")
        if "A11" not in area and not area.startswith("Bill!A1:"):
            bad("Bill print area does not include the line table")
        refs = {str(dv.sqref) for dv in bl.data_validations.dataValidation}
        for want in ("O2", "O9", "O10"):
            if want not in refs:
                bad(f"Bill: no dropdown on {want}")
    gp = [ws_g for ws_g in wb["GRep"].iter_rows(min_row=13, max_row=24, min_col=3, max_col=3,
                                                values_only=True)]
    print("   GRep formula sample:", str(gp[0][0])[:90])
    if "Book!$O$10:$O$50009" not in str(gp[0][0]):
        bad("GRep does not sum the full Book range O10:O50009")
    print("   Inv print area:", wb["Inv"].print_area, "| Print:", wb["Print"].print_area)
    wb.close()


def bank_structure():
    print("\n02_Bank.xlsx — bank ledger and payable")
    p = os.path.join(OUT, "02_Bank.xlsx")
    wb = openpyxl.load_workbook(p, data_only=False)
    txn = wb["Txn"]
    rows = 0
    bd = Counter()
    for r in range(13, txn.max_row + 1):
        amt = txn.cell(r, 13).value                # M = مبلغ
        if not isinstance(amt, (int, float)):
            continue
        rows += 1
        op = str(txn.cell(r, 5).value or "")        # E = جزئیات عملیات
        person = str(txn.cell(r, 17).value or "")   # Q = شخص باسکول دیگران
        if "باسکول دیگران" in op:
            bd[person] += 1
    post = wb["Post"]
    pay = wb["Payable"]
    print(f"   Txn filled rows: {rows}   Post dims: {post.dimensions}   "
          f"Payable: {pay.dimensions}")
    if txn.max_row < 5000:
        bad(f"Txn capacity only {txn.max_row} rows")
    if post.max_row < 10000:
        bad(f"Post rows only {post.max_row}")
    import zipfile as _zf
    with _zf.ZipFile(p) as z:
        rel = {}
        for n in z.namelist():
            if n.startswith("xl/worksheets/sheet"):
                t = z.read(n).decode()
                s_ = re.findall(r'<dataValidation [^>]*sqref="([^"]+)"', t)
                rel[n] = s_
    bank_dv = [x for v in rel.values() for x in v if "5012" in x or "5000" in x]
    print("   Txn dropdowns reach row 5012:", bool(bank_dv))
    people = 0
    for r in range(13, 60):
        code = pay.cell(r, 2).value
        if isinstance(code, str) and code.startswith("BD"):
            people += 1
    print(f"   Payable persons: {people}")
    if people < 5:
        bad("Payable report lists too few people")
    print("   Payable sample:", str(pay.cell(13, 3).value)[:80])
    for name in ("Anal", "Class", "Personal"):
        if name not in wb.sheetnames:
            bad(f"02_Bank: no {name} sheet")
            continue
        ws2 = wb[name]
        print(f"   {name}: dims {ws2.dimensions}  print {ws2.print_area}  "
              f"charts {len(ws2._charts)}")
        if not ws2.print_area:
            bad(f"02_Bank/{name}: no print area")
        if not ws2._charts and name != "Personal":
            bad(f"02_Bank/{name}: no chart")
    # bank txn rows that name a باسکول دیگران person
    print("   باسکول دیگران rows in Txn:", sum(bd.values()), dict(bd))
    if not sum(bd.values()):
        bad("no «باسکول دیگران» transaction in 02_Bank → Payable report is empty")
    wb.close()


def cross_links():
    print("\nCross-file links (Home → other files) and Sync coverage")
    import glob
    for p in sorted(glob.glob(os.path.join(OUT, "*.xlsx"))):
        wb = openpyxl.load_workbook(p, data_only=False)
        f = os.path.basename(p)
        if "Sync" not in wb.sheetnames:
            bad(f"{f}: no Sync sheet")
        if "Help" not in wb.sheetnames:
            bad(f"{f}: no Help sheet")
        ext = []
        for ws in wb.worksheets:
            for c in list(getattr(ws, "_cells", {}).values()):
                h = getattr(c, "hyperlink", None)
                t = getattr(h, "target", "") if h is not None else ""
                if t and ".xlsx" in t:
                    ext.append(t)
        with zipfile.ZipFile(p) as zf:
            parts = [n for n in zf.namelist() if n.startswith("xl/externalLinks/externalLink")]
        files = sorted({t.split("#")[0] for t in ext})
        print(f"   {f:<18} sync={('Sync' in wb.sheetnames)!s:<5} "
              f"help={('Help' in wb.sheetnames)!s:<5} "
              f"live links={len(parts)} file links={len(files)} {files[:3]}")
        if not parts and not files and f not in ("07_Demo.xlsx",):
            bad(f"{f}: no live link to another file")
        wb.close()


def mirrors():
    print("\nMirror files — 50,009-row data blocks")
    for name in ("03_Report.xlsx", "04_Inv.xlsx", "05_Bal.xlsx", "06_Buyers.xlsx"):
        p = os.path.join(OUT, name)
        wb = openpyxl.load_workbook(p, read_only=True)
        ws = wb["Data"]
        n = 0
        for row in ws.iter_rows(min_row=10, max_row=50009, max_col=2, values_only=True):
            if row[1] not in (None, ""):
                n += 1
        print(f"   {name:<16} Data max_row {ws.max_row:<7} filled {n}")
        if ws.max_row < 50009:
            bad(f"{name}: Data only {ws.max_row} rows")
        wb.close()


def main(verbose=False):
    files = sorted(f for f in os.listdir(OUT) if f.endswith(".xlsx"))
    print(f"v04 files: {len(files)}")
    if len(files) != 9:
        bad(f"expected 9 workbooks, found {len(files)}")
    for f in files:
        audit_file(os.path.join(OUT, f), verbose)
    book_structure()
    bank_structure()
    cross_links()
    mirrors()
    print("\n" + "=" * 70)
    print(f"issues: {len(issues)}")
    for i in issues:
        print("   •", i)
    return 0 if not issues else 1


if __name__ == "__main__":
    sys.exit(main("--verbose" in sys.argv))
