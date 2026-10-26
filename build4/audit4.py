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
    if bk.max_row != 60010:
        bad(f"Book last row {bk.max_row} ≠ 60010")
    # sampling the last row of every block
    want = {10: "Scale", 40009: "Scale", 40010: "Pay", 55009: "Pay",
            55010: "Crate", 60009: "Crate"}
    seen = {}
    for r, row in enumerate(bk.iter_rows(min_row=1, max_row=60010, values_only=True), 1):
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
                        ("Book", "Q60010")):
        area = str(wb[sheet].print_area or "").replace("$", "")
        if want not in area:
            bad(f"{sheet} print area {area} does not reach {want}")
    for sheet in ("Inv", "Print", "GRep"):
        ws = wb[sheet]
        print(f"   {sheet}: dims {ws.dimensions}  print {ws.print_area}  "
              f"dv {len(ws.data_validations.dataValidation)}")
    for bname, pt in (("Bill", "دریافت"), ("BillBuy", "پرداخت")):
        if bname not in wb.sheetnames:
            bad(f"01_Entry: no {bname} sheet")
            continue
        bl = wb[bname]
        area = str(bl.print_area or "").replace("$", "").replace("'", "")
        net = str(bl["H11"].value or "")
        print(f"   {bname}: dims {bl.dimensions}  print {area}  "
              f"dv {len(bl.data_validations.dataValidation)}")
        if not area.startswith(f"{bname}!A1:K"):
            bad(f"{bname} print area {area} is not A1:K…")
        if "FLOOR(" not in net or "$O$5" not in net:
            bad(f"{bname} net-kg formula missing FLOOR / loss cell")
        pay = ""
        for r_ in range(30, bl.max_row + 1):
            v = str(bl.cell(r_, 4).value or "")
            if "SUMIFS(Pay!" in v:
                pay = v
                break
        if pt not in pay:
            bad(f"{bname} payment total does not filter on «{pt}»")
        refs = {str(dv.sqref) for dv in bl.data_validations.dataValidation}
        for want in ("O2", "O9", "O10"):
            if want not in refs:
                bad(f"{bname}: no dropdown on {want}")
        pics = len(getattr(bl, "_images", []))
        head = str(bl["C1"].value or "")
        fit_h = str(bl.page_setup.fitToHeight)
        print(f"   {bname} letterhead: name={head[:28]!r} images={pics} fitToHeight={fit_h}")
        if pics < 2:
            bad(f"{bname}: letterhead / signature seal missing ({pics} images)")
        if not head.strip():
            bad(f"{bname}: company name in the letterhead is empty")
        if fit_h != "1":
            bad(f"{bname}: fitToHeight={fit_h} (the invoice must fit one A4 page)")
    gp = [ws_g for ws_g in wb["GRep"].iter_rows(min_row=13, max_row=24, min_col=3, max_col=3,
                                                values_only=True)]
    print("   GRep formula sample:", str(gp[0][0])[:90])
    if "Book!$O$10:$O$60009" not in str(gp[0][0]):
        bad("GRep does not sum the full Book range O10:O60009")
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
    dn = {k: v.attr_text for k, v in wb.defined_names.items()}
    grp = [k for k in dn if k.startswith("ClsG")]
    print(f"   class groups: {len(grp)} ranges · ListClsParent={dn.get('ListClsParent')}")
    if len(grp) < 5 or "ClsMap" not in dn or "ListClsParent" not in dn:
        bad("02_Bank: dependent category dropdowns are not wired (ClsG*/ClsMap)")
    dv_src = [str(dv.formula1) for dv in txn.data_validations.dataValidation]
    if not any("INDIRECT" in x for x in dv_src):
        bad("02_Bank/Txn: sub-category dropdown is not dependent (no INDIRECT)")
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


def master_cls():
    print("\n00_Master.xlsx — Cls (two-column classification)")
    p = os.path.join(OUT, "00_Master.xlsx")
    wb = openpyxl.load_workbook(p, read_only=True)
    if "Cls" not in wb.sheetnames:
        bad("00_Master: no Cls sheet")
        wb.close()
        return
    ws = wb["Cls"]
    heads = [ws.cell(8, c) for c in range(1, 8)]
    print("   heads:", [h.value for h in heads if h.value])
    n = 0
    parents = set()
    for r in range(10, 400):
        row = [ws.cell(r, c).value for c in range(1, 8)]
        if not row[2]:
            continue
        n += 1
        parents.add(row[1])
    print(f"   rows: {n}  parents: {len(parents)}  {sorted(parents)[:3]}")
    if n < 20 or len(parents) < 5:
        bad("00_Master/Cls: classification table too small")
    wb.close()


def dashboard_cls():
    print("\n08_Dash.xlsx — Cls (category ↔ sub-category report)")
    p = os.path.join(OUT, "08_Dash.xlsx")
    wb = openpyxl.load_workbook(p, data_only=False)
    if "Cls" not in wb.sheetnames:
        bad("08_Dash: no Cls sheet")
        wb.close()
        return
    ws = wb["Cls"]
    print(f"   dims {ws.dimensions}  print {ws.print_area}  charts {len(ws._charts)}")
    parents = kids = 0
    for r in range(14, ws.max_row + 1):
        b = str(ws.cell(r, 2).value or "")
        if b == "جمع طبقه":
            parents += 1
        elif "└" in b:
            kids += 1
    print(f"   parents {parents}  sub-categories {kids}")
    if parents < 5 or kids < 20:
        bad(f"08_Dash/Cls: only {parents} parents / {kids} sub-categories")
    if not ws._charts:
        bad("08_Dash/Cls: no chart")
    ref = str(ws.cell(15, 3).value or "")
    if "F2!" not in ref:
        bad("08_Dash/Cls: the table does not read the bank feed (F2)")
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
    print("\nMirror files — 60,009-row data blocks")
    for name in ("03_Report.xlsx", "04_Inv.xlsx", "05_Bal.xlsx", "06_Buyers.xlsx"):
        p = os.path.join(OUT, name)
        wb = openpyxl.load_workbook(p, read_only=True)
        ws = wb["Data"]
        n = 0
        for row in ws.iter_rows(min_row=10, max_row=60009, max_col=2, values_only=True):
            if row[1] not in (None, ""):
                n += 1
        print(f"   {name:<16} Data max_row {ws.max_row:<7} filled {n}")
        if ws.max_row < 60009:
            bad(f"{name}: Data only {ws.max_row} rows")
        wb.close()


def partnership():
    """شراکت — the partner register, the share columns and the partner report."""
    print("\nPartnership — شرکا، سهم وزنی و گزارش Shrk")
    wb = openpyxl.load_workbook(os.path.join(OUT, "00_Master.xlsx"))
    ok = True
    if "Shrk" not in wb.sheetnames:
        bad("00_Master: sheet Shrk is missing"); ok = False
    else:
        ws = wb["Shrk"]
        codes = [ws.cell(r, 1).value for r in range(12, 32)]
        names = [ws.cell(r, 2).value for r in range(12, 32)]
        print(f"   Shrk: {sum(1 for n in names if n)} named partners · code SH01…SH20 "
              f"· print {ws.print_area}")
        if codes[0] != "SH01" or codes[-1] != "SH20":
            bad("00_Master/Shrk: partner codes are not SH01…SH20"); ok = False
        if not ws.print_area:
            bad("00_Master/Shrk: no print area"); ok = False
        if len(ws.data_validations.dataValidation) < 3:
            bad("00_Master/Shrk: dropdowns are missing"); ok = False
    names = {n: d.attr_text for n, d in wb.defined_names.items()}
    for n in ("ShrkTbl", "ListShrk"):
        if n not in names:
            bad(f"00_Master: name {n} is missing"); ok = False
    acc = wb["Acc"]
    sh = [acc.cell(r, 1).value for r in range(255, 285)]
    if "SH01" not in sh or "SH20" not in sh:
        bad("00_Master/Acc: partner accounts SH01…SH20 are missing"); ok = False
    op = wb["Op"]
    ops = [(op.cell(r, 1).value, op.cell(r, 2).value) for r in range(10, 34)]
    if not any("شراکتی" in str(t) for _, t in ops):
        bad("00_Master/Op: partnership operations are missing"); ok = False
    else:
        print("   Op:", [t for _, t in ops if "شراکتی" in str(t)])
    wb.close()

    wb = openpyxl.load_workbook(os.path.join(OUT, "01_Entry.xlsx"))
    sc = wb["Scale"]
    heads = [sc[f"{c}8"].value for c in ("AQ", "AR", "AS", "AT", "AU", "AV", "AW")]
    if any(h is None for h in heads):
        bad("01_Entry/Scale: share headers AQ…AV are incomplete"); ok = False
    else:
        print("   Scale AQ…AV:", " · ".join(heads))
    if not any("AQ" in str(dv.sqref) for dv in sc.data_validations.dataValidation):
        bad("01_Entry/Scale: AQ has no partner dropdown"); ok = False
    if str(sc.print_area) != "'Scale'!$A$1:$T$10009":
        bad(f"01_Entry/Scale: print area is {sc.print_area}"); ok = False
    last = 10009
    for col in ("AS", "AT", "AU", "AV"):          # our / partner share, rial and kg
        f = str(sc[f"{col}{last}"].value or "")
        if not (f.startswith("=") and "$AR" in f):
            bad(f"01_Entry/Scale: {col}{last} is not a share formula: {f[:60]}")
            ok = False
    if "SEARCH" not in str(sc[f"AW{last}"].value or ""):
        bad(f"01_Entry/Scale: AW{last} has no partnership control"); ok = False
    if "Shrk" not in wb.sheetnames:
        bad("01_Entry: sheet Shrk (partner report) is missing"); ok = False
    else:
        ws = wb["Shrk"]
        print(f"   Shrk report: dims {ws.dimensions} · print {ws.print_area}")
        if str(ws.print_area) != "'Shrk'!$A$1:$N$189":
            bad(f"01_Entry/Shrk: print area is {ws.print_area}"); ok = False
        if str(ws["C9"].value) != "خرید شراکتی (ورود)":
            bad("01_Entry/Shrk: C9 must hold the partnership-in operation"); ok = False
        if "MATCH" not in str(ws["O38"].value or ""):
            bad("01_Entry/Shrk: detail block has no row lookup"); ok = False
    names = {n: d.attr_text for n, d in wb.defined_names.items()}
    if names.get("ShrkTbl") != "Sync!$CS$10:$CV$29":
        bad(f"01_Entry: ShrkTbl is {names.get('ShrkTbl')}"); ok = False
    if "280" not in str(names.get("AcctTbl")):
        bad(f"01_Entry: AcctTbl does not reach the partners ({names.get('AcctTbl')})")
    wb.close()

    wb = openpyxl.load_workbook(os.path.join(OUT, "02_Bank.xlsx"))
    n = len(wb.defined_names["ListTxnOps"].attr_text.split("$")[-1].split(":")[-1])
    txt = wb.defined_names["ListTxnOps"].attr_text
    print(f"   Bank ops table: {txt}")
    wb.close()
    return ok


def main(verbose=False):
    files = sorted(f for f in os.listdir(OUT) if f.endswith(".xlsx"))
    print(f"v04 files: {len(files)}")
    if len(files) != 9:
        bad(f"expected 9 workbooks, found {len(files)}")
    for f in files:
        audit_file(os.path.join(OUT, f), verbose)
    book_structure()
    master_cls()
    bank_structure()
    dashboard_cls()
    partnership()
    cross_links()
    mirrors()
    print("\n" + "=" * 70)
    print(f"issues: {len(issues)}")
    for i in issues:
        print("   •", i)
    return 0 if not issues else 1


if __name__ == "__main__":
    sys.exit(main("--verbose" in sys.argv))
