# -*- coding: utf-8 -*-
"""Accounting audit of the v03 series — reads the harvested numbers, no Excel.

    python build2/audit3.py [--verbose]

Checks
  01  the journal balances (total, every document, every account code known)
  02  no empty/zero lines in the journal
  03  all bank legs go through the clearing account X011
  04  the bank sub-ledger balances and the clearing nets to zero
  05  bank balances = opening + debits − credits (per bank)
  06  no bank code is used as a ledger account in the journal
"""
import json
import os
import sys
from collections import defaultdict

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
OUT = os.path.join(ROOT, "Mazandasht_v03")
CACHE = os.path.join(HERE, "feed_cache.json")
CLEARING = "X011"
POS_BANK = "BN02"

OK, BAD = "✓", "✗"


def load():
    with open(CACHE, encoding="utf-8") as fh:
        data = json.load(fh)

    def rows(prefix):
        out = defaultdict(dict)
        for k, v in data.get(prefix, {}).items():
            col, ref = k.split("|")
            out[int(ref[1:])][col] = v
        return out

    return data, rows("F1"), rows("F2")


def master_codes():
    wb = openpyxl.load_workbook(os.path.join(OUT, "00_Master.xlsx"), read_only=True)
    acc = wb["Acc"]
    codes = {}
    for r in range(10, 260):
        c = acc.cell(r, 1).value
        if c:
            codes[str(c)] = acc.cell(r, 4).value
    banks, opening = {}, {}
    bnk = wb["Bnk"]
    for r in range(10, 20):
        c = bnk.cell(r, 1).value
        if c:
            banks[str(c)] = bnk.cell(r, 2).value
            opening[str(c)] = float(bnk.cell(r, 3).value or 0)
    wb.close()
    return codes, banks, opening


def main(verbose=False):
    data, book, post = load()
    codes, banks, opening = master_codes()
    issues = []

    # ---------------------------------------------------------- journal
    tot_dr = sum(float(r.get("L") or 0) for r in book.values())
    tot_cr = sum(float(r.get("M") or 0) for r in book.values())
    print(f"01  journal rows            {len(book):>5}")
    print(f"    debit / credit          {tot_dr:,.0f} / {tot_cr:,.0f}   "
          f"{OK if round(tot_dr - tot_cr) == 0 else BAD}")
    if round(tot_dr - tot_cr):
        issues.append("journal not balanced")

    docs = defaultdict(lambda: [0.0, 0.0])
    for r in book.values():
        d = r.get("B") or ""
        docs[d][0] += float(r.get("L") or 0)
        docs[d][1] += float(r.get("M") or 0)
    unbal = {k: v for k, v in docs.items() if round(v[0] - v[1])}
    print(f"02  documents               {len(docs):>5}   unbalanced {len(unbal)} "
          f"{OK if not unbal else BAD}")
    if unbal:
        issues.append(f"unbalanced docs: {list(unbal)[:3]}")

    zero = [r for r in book.values()
            if not str(r.get("D") or "").strip() or (not (r.get("L") or r.get("M")))]
    print(f"03  empty / zero lines      {len(zero):>5}   {OK if not zero else BAD}")
    if zero:
        issues.append(f"{len(zero)} empty journal lines")

    unknown = {str(r.get("D")) for r in book.values()
               if str(r.get("D") or "") not in codes}
    print(f"04  unknown account codes   {len(unknown):>5}   {OK if not unknown else BAD}"
          f"   {sorted(unknown)[:5] if unknown else ''}")
    if unknown:
        issues.append(f"unknown codes {sorted(unknown)[:5]}")

    bank_ledger = {str(r.get("D")) for r in book.values() if str(r.get("D")) in banks}
    print(f"05  bank codes in journal   {len(bank_ledger):>5}   "
          f"{OK if not bank_ledger else BAD}   {sorted(bank_ledger) if bank_ledger else ''}")
    if bank_ledger:
        issues.append("journal posts directly to a bank account")

    clearing = [r for r in book.values() if str(r.get("D")) == CLEARING]
    cl_dr = sum(float(r.get("L") or 0) for r in clearing)
    cl_cr = sum(float(r.get("M") or 0) for r in clearing)
    print(f"06  clearing {CLEARING} rows      {len(clearing):>5}   "
          f"dr {cl_dr:,.0f} / cr {cl_cr:,.0f}")

    # ------------------------------------------------------------- bank
    b_dr = sum(float(r.get("I") or 0) for r in post.values())
    b_cr = sum(float(r.get("J") or 0) for r in post.values())
    print(f"07  bank lines              {len(post):>5}   "
          f"dr {b_dr:,.0f} / cr {b_cr:,.0f}   "
          f"{OK if round(b_dr - b_cr) == 0 else BAD}")
    if round(b_dr - b_cr):
        issues.append("bank ledger not balanced")

    b_cl = [r for r in post.values() if str(r.get("D")) == CLEARING]
    b_cl_dr = sum(float(r.get("I") or 0) for r in b_cl)
    b_cl_cr = sum(float(r.get("J") or 0) for r in b_cl)
    net = (cl_dr - cl_cr) + (b_cl_dr - b_cl_cr)
    print(f"08  clearing net (journal + bank)   {net:,.0f}   "
          f"{OK if round(net) == 0 else BAD}")
    if round(net):
        issues.append(f"clearing {CLEARING} does not net to zero ({net:,.0f})")

    bad_bank = {str(r.get("D")) for r in post.values()
                if str(r.get("D")) not in codes}
    print(f"09  unknown codes in bank   {len(bad_bank):>5}   "
          f"{OK if not bad_bank else BAD}  {sorted(bad_bank)[:5] if bad_bank else ''}")
    if bad_bank:
        issues.append(f"unknown bank codes {sorted(bad_bank)[:5]}")

    # ------------------------------------------------------- POS control
    j_pos = sum(float(r.get("L") or 0) for r in book.values()
                if str(r.get("D")) == CLEARING and str(r.get("N") or "").strip() == "پوز")
    b_pos = sum(float(r.get("I") or 0) for r in post.values()
                if str(r.get("D") or "").startswith("BN")
                and str(r.get("G") or "") == "واریز پز باسکول (۲۰۰۶)")
    print(f"10  POS control   journal {j_pos:>15,.0f}   bank {b_pos:>15,.0f}   "
          f"{OK if round(j_pos - b_pos) == 0 else BAD}")
    if round(j_pos - b_pos):
        issues.append(f"POS journal {j_pos:,.0f} ≠ bank {b_pos:,.0f}")

    # every clearing line must carry a bank code (otherwise it lands on BN02)
    no_bank = [r for r in book.values()
               if str(r.get("D")) == CLEARING and not str(r.get("P") or "").strip()]
    print(f"11  clearing lines w/o bank   {len(no_bank):>5}   "
          f"{OK if not no_bank else BAD}")
    if no_bank:
        issues.append(f"{len(no_bank)} clearing lines without a bank code")

    # ------------------------------------------------- per-bank balances
    move = defaultdict(lambda: [0.0, 0.0])
    for r in post.values():
        c = str(r.get("D") or "")
        if c in banks:
            move[c][0] += float(r.get("I") or 0)
            move[c][1] += float(r.get("J") or 0)
    total = 0.0
    print("\n    bank                 opening            debit           credit"
          "          balance")
    for code, title in banks.items():
        dr, cr = move[code]
        bal = opening[code] + dr - cr
        total += bal
        if dr or cr or opening[code]:
            print(f"    {code}  {str(title)[:16]:<16} {opening[code]:>14,.0f} "
                  f"{dr:>15,.0f} {cr:>15,.0f} {bal:>15,.0f}")
    print(f"    {'total':<35} {total:>60,.0f}")

    # ---------------------------------------------------------- summary
    print(f"\nissues: {len(issues)}")
    for i in issues:
        print("   •", i)
    if verbose:
        print("\njournal accounts:")
        agg = defaultdict(lambda: [0.0, 0.0, 0])
        for r in book.values():
            c = str(r.get("D"))
            agg[c][0] += float(r.get("L") or 0)
            agg[c][1] += float(r.get("M") or 0)
            agg[c][2] += 1
        for c in sorted(agg):
            v = agg[c]
            print(f"   {c:8s} rows={v[2]:4d} dr={v[0]:>15,.0f} cr={v[1]:>15,.0f} "
                  f"bal={v[0]-v[1]:>15,.0f}  {codes.get(c, '???')}")
    return 0 if not issues else 1


if __name__ == "__main__":
    sys.exit(main("--verbose" in sys.argv))
