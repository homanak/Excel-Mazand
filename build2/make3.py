# -*- coding: utf-8 -*-
"""Builds the whole Mazandasht v03 series.

    python build2/make3.py            # full build (harvests dashboard feed)
    python build2/make3.py --fast     # skip value harvesting (links show 0)
"""
import json
import os
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)

import kit          # noqa: E402
import unify        # noqa: E402
import bank2        # noqa: E402
import dash         # noqa: E402
import links        # noqa: E402
import master       # noqa: E402

V02 = os.path.join(ROOT, "Mazandasht_v02")
OUT = os.path.join(ROOT, "Mazandasht_v03")
CACHE = os.path.join(HERE, "feed_cache.json")


def patch_entry():
    """Code unification + clearing account + zero-guard on 01_Entry."""
    entry = os.path.join(OUT, "01_Entry.xlsx")
    hits = unify.unify_codes(entry)
    ops = unify.entry_clearing(entry)
    unify.entry_pos(entry)
    guard = unify.entry_zero_guard(entry)
    print(f"  01_Entry: {hits} code cells, {ops} clearing rules, {guard} zero guards")
    for name in ("03_Report.xlsx", "04_Inv.xlsx", "05_Bal.xlsx", "06_Buyers.xlsx",
                 "07_Demo.xlsx"):
        n = unify.unify_codes(os.path.join(OUT, name))
        print(f"  {name}: {n} code cells")


def statement_rows(book):
    """Bank-statement lines implied by the journal's clearing rows (X011).

    The journal posts every bank leg to the clearing account X011; the real
    money then shows up on the bank statement.  The clearing movement is
    grouped per day, per bank and per channel — POS takings, other receipts and
    payments — and each group becomes one statement line, so the clearing
    account closes to zero once the statement is complete.
    """
    net = {}
    for ref, data in book.items():
        if data.get("D") != unify.CLEARING:
            continue
        dr, cr = data.get("L") or 0, data.get("M") or 0
        if dr <= 0 and cr <= 0:
            continue
        day = data.get("A") or ""
        bank = data.get("P") or unify.POS_BANK
        op = str(data.get("N") or "")
        if dr > 0:
            chan = "pos" if op.strip() == "پوز" else "other"
        else:
            chan = "pay"
        key = (day, bank, chan)
        agg = net.setdefault(key, {"dr": 0.0, "cr": 0.0, "docs": [], "desc": ""})
        agg["dr"] += float(dr)
        agg["cr"] += float(cr)
        d = str(data.get("B") or "")
        if d and d not in agg["docs"]:
            agg["docs"].append(d)
        if not agg["desc"] and data.get("E"):
            agg["desc"] = str(data.get("E"))[:40]

    rows = []
    for (day, bank, chan), agg in sorted(net.items()):
        diff = agg["dr"] - agg["cr"]
        if round(diff) == 0:
            continue
        ref = "JB-" + "/".join(agg["docs"][:2])
        if diff > 0:
            if chan == "pos":
                op, method = "واریز پز باسکول (۲۰۰۶)", "کارت/پوز"
            else:
                op, method = "تسویه دریافت (واسط)", "انتقال بانکی"
            rows.append((day, ref, op, bank, "", method, "",
                         agg["desc"] or "واریز از واسط", ref, diff))
        else:
            rows.append((day, ref, "تسویه پرداخت (واسط)", bank, "", "انتقال بانکی",
                         "", agg["desc"] or "پرداخت از واسط", ref, -diff))
    return rows


def copy_set():
    for old, new in (("01_Entry.xlsx", "01_Entry.xlsx"), ("03_Report.xlsx", "03_Report.xlsx"),
                     ("04_Inv.xlsx", "04_Inv.xlsx"), ("05_Bal.xlsx", "05_Bal.xlsx"),
                     ("06_Buyers.xlsx", "06_Buyers.xlsx"), ("07_Demo.xlsx", "07_Demo.xlsx")):
        src = os.path.join(V02, old)
        if os.path.exists(src):
            shutil.copy2(src, os.path.join(OUT, new))


def feed_values(fast=False):
    """Harvest the real numbers once per build (cached in feed_cache.json)."""
    if fast:
        return {}
    if os.path.exists(CACHE):
        with open(CACHE, encoding="utf-8") as fh:
            data = json.load(fh)
        return {"dates": data.get("dates") or [],
                "ops": data.get("ops") or [],
                "statement": data.get("statement") or [],
                "F1": {tuple(k.split("|")): v for k, v in data.get("F1", {}).items()},
                "F2": {}}          # Post is always re-read after 02_Bank is built

    print("  harvesting Book values (01_Entry) …", flush=True)
    cells = links.harvest_cells("Book", dash.BOOK_COLS + ["P"], 10, 1808)
    f1raw, m1 = links.harvest(os.path.join(OUT, "01_Entry.xlsx"), {"Book": cells},
                              os.path.join(HERE, "_h1.json"))

    f1 = {}
    for col in dash.BOOK_COLS + ["P"]:
        for r in range(10, 1809):
            v = f1raw["Book"].get(f"{col}{r}")
            if v not in (None, "", 0):
                f1[(col, f"{col}{r}")] = v

    book = {}
    for r in range(10, 1810):
        row = {c: f1raw["Book"].get(f"{c}{r}") for c in dash.BOOK_COLS + ["P"]}
        if row.get("A") in (None, "", 0):
            continue
        book[f"r{r}"] = row
    stmt = statement_rows(book)

    dates, seen = [], set()
    for (col, ref), v in f1.items():
        if col == "A" and isinstance(v, str) and v and v not in seen:
            seen.add(v)
            dates.append(v)
    dates.sort()

    pack = {"dates": dates,
            "ops": master.read_ops(os.path.join(OUT, "00_Master.xlsx")),
            "statement": stmt,
            "F1": {f"{k[0]}|{k[1]}": v for k, v in f1.items()},
            "F2": {}}
    with open(CACHE, "w", encoding="utf-8") as fh:
        json.dump(pack, fh, ensure_ascii=False)
    print(f"  book: {len(f1)} cells, {len(stmt)} statement rows (miss {m1})")
    return {"dates": dates, "ops": pack["ops"], "statement": stmt, "F1": f1, "F2": {}}


def harvest_post(vals):
    """Read 02_Bank!Post *after* the bank file exists."""
    print("  harvesting Post values (02_Bank) …", flush=True)
    cells = links.harvest_cells("Post", dash.POST_SRC, 13, 292)
    f2raw, m2 = links.harvest(os.path.join(OUT, "02_Bank.xlsx"), {"Post": cells},
                              os.path.join(HERE, "_h2.json"))
    f2 = {}
    for col in dash.POST_SRC:
        for r in range(13, 293):
            v = f2raw["Post"].get(f"{col}{r}")
            if v not in (None, "", 0):
                f2[(col, f"{col}{r}")] = v
    vals["F2"] = f2
    vals["post_raw"] = f2raw["Post"]
    if os.path.exists(CACHE):
        with open(CACHE, encoding="utf-8") as fh:
            data = json.load(fh)
        data["F2"] = {f"{k[0]}|{k[1]}": v for k, v in f2.items()}
        with open(CACHE, "w", encoding="utf-8") as fh:
            json.dump(data, fh, ensure_ascii=False)
    print(f"  post: {len(f2)} cells (miss {m2})")
    return vals


def main(fast=False):
    os.makedirs(OUT, exist_ok=True)
    kit.ICONS = kit.make_icons()

    print("00_Master.xlsx")
    master.build(os.path.join(V02, "00_Master.xlsx"), os.path.join(OUT, "00_Master.xlsx"))

    print("01_Entry …")
    copy_set()
    patch_entry()

    print("feed — journal")
    vals = feed_values(fast)

    print("02_Bank.xlsx")
    demo = [tuple(r) for r in vals.get("statement") or []]
    demo += [
        ("1405/07/13", "TB-901", "انتقال داخلی بین بانک‌ها", "BN01", "BN03",
         "انتقال بانکی", "", "انتقال به حساب ملت", "TR-9001", 50000000),

    ]
    bank2.build(os.path.join(OUT, "00_Master.xlsx"), os.path.join(OUT, "02_Bank.xlsx"),
                demo=demo or None)
    print(f"  statement rows: {len(demo)}")

    print("05_Bal IBank → 02_Bank")
    unify.bal_ibank(os.path.join(OUT, "05_Bal.xlsx"))
    if vals.get("F2") is None or not vals.get("F2"):
        harvest_post(vals)
    post_vals = {ref: v for (c, ref), v in vals.get("F2", {}).items()}
    links.patch(os.path.join(OUT, "05_Bal.xlsx"),
                [{"file": "02_Bank.xlsx", "sheets": ["Post"],
                  "values": {"Post": post_vals}}])

    print("08_Dash.xlsx")
    tmp = os.path.join(HERE, "_dash.xlsx")
    dash.build(os.path.join(OUT, "00_Master.xlsx"), os.path.join(OUT, "02_Bank.xlsx"),
               os.path.join(OUT, "01_Entry.xlsx"), tmp, vals)
    links.patch(tmp, [
        {"file": "01_Entry.xlsx", "sheets": ["Book"],
         "values": {"Book": {ref: v for (c, ref), v in vals.get("F1", {}).items()}}},
        {"file": "02_Bank.xlsx", "sheets": ["Post"],
         "values": {"Post": post_vals}},
    ])
    shutil.move(tmp, os.path.join(OUT, "08_Dash.xlsx"))
    print("  -> 08_Dash.xlsx", os.path.getsize(os.path.join(OUT, "08_Dash.xlsx")) // 1024, "KB")
    print("done")


if __name__ == "__main__":
    main("--fast" in sys.argv)
