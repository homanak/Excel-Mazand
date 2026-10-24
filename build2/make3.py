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
import bank2        # noqa: E402
import dash         # noqa: E402
import links        # noqa: E402
import master       # noqa: E402

V02 = os.path.join(ROOT, "Mazandasht_v02")
OUT = os.path.join(ROOT, "Mazandasht_v03")
CACHE = os.path.join(HERE, "feed_cache.json")


def copy_set():
    for old, new in (("01_Entry.xlsx", "01_Entry.xlsx"), ("03_Report.xlsx", "03_Report.xlsx"),
                     ("04_Inv.xlsx", "04_Inv.xlsx"), ("05_Bal.xlsx", "05_Bal.xlsx"),
                     ("06_Buyers.xlsx", "06_Buyers.xlsx"), ("07_Demo.xlsx", "07_Demo.xlsx")):
        src = os.path.join(V02, old)
        if os.path.exists(src):
            shutil.copy2(src, os.path.join(OUT, new))


def feed_values(fast=False):
    if os.path.exists(CACHE):
        with open(CACHE, encoding="utf-8") as fh:
            data = json.load(fh)
        dates = data.get("dates") or []
        ops = data.get("ops") or []
        f1 = {tuple(k.split("|")): v for k, v in data.get("F1", {}).items()}
        f2 = {tuple(k.split("|")): v for k, v in data.get("F2", {}).items()}
        return {"dates": dates, "ops": ops, "F1": f1, "F2": f2}

    if fast:
        return {}
    entry = os.path.join(OUT, "01_Entry.xlsx")
    bank = os.path.join(OUT, "02_Bank.xlsx")
    print("  harvesting Book values (01_Entry) …", flush=True)
    cells = links.harvest_cells("Book", dash.BOOK_COLS, 10, 1808)
    f1raw, m1 = links.harvest(entry, {"Book": cells}, os.path.join(HERE, "_h1.json"))
    print("  harvesting Post values (02_Bank) …", flush=True)
    cells2 = links.harvest_cells("Post", dash.POST_SRC, 13, 292)
    f2raw, m2 = links.harvest(bank, {"Post": cells2}, os.path.join(HERE, "_h2.json"))

    f1 = {}
    for col in dash.BOOK_COLS:
        for r in range(10, 1809):
            v = f1raw["Book"].get(f"{col}{r}")
            if v not in (None, "", 0):
                f1[(col, f"{col}{r}")] = v
    f2 = {}
    for col in dash.POST_SRC:
        for r in range(13, 293):
            v = f2raw["Post"].get(f"{col}{r}")
            if v not in (None, "", 0):
                f2[(col, f"{col}{r}")] = v

    dates, seen = [], set()
    for (col, ref), v in f1.items():
        if col == "A" and isinstance(v, str) and v and v not in seen:
            seen.add(v)
            dates.append(v)
    for (col, ref), v in f2.items():
        if col == "A" and isinstance(v, str) and v and v not in seen:
            seen.add(v)
            dates.append(v)
    dates.sort()

    pack = {"dates": dates,
            "ops": master.read_ops(os.path.join(OUT, "00_Master.xlsx")),
            "F1": {f"{k[0]}|{k[1]}": v for k, v in f1.items()},
            "F2": {f"{k[0]}|{k[1]}": v for k, v in f2.items()}}
    with open(CACHE, "w", encoding="utf-8") as fh:
        json.dump(pack, fh, ensure_ascii=False)
    print(f"  feed: {len(f1)} F1 cells, {len(f2)} F2 cells, {len(dates)} dates "
          f"(miss {m1}/{m2})")
    return {"dates": dates, "ops": pack["ops"], "F1": f1, "F2": f2}


def main(fast=False):
    os.makedirs(OUT, exist_ok=True)
    kit.ICONS = kit.make_icons()

    print("00_Master.xlsx")
    master.build(os.path.join(V02, "00_Master.xlsx"), os.path.join(OUT, "00_Master.xlsx"))

    print("01_Entry …")
    copy_set()

    print("02_Bank.xlsx")
    bank2.build(os.path.join(OUT, "00_Master.xlsx"), os.path.join(OUT, "02_Bank.xlsx"))

    print("08_Dash.xlsx")
    vals = feed_values(fast)
    tmp = os.path.join(HERE, "_dash.xlsx")
    dash.build(os.path.join(OUT, "00_Master.xlsx"), os.path.join(OUT, "02_Bank.xlsx"),
               os.path.join(OUT, "01_Entry.xlsx"), tmp, vals)
    links.patch(tmp, [
        {"file": "01_Entry.xlsx", "sheets": ["Book"],
         "values": {"Book": {ref: v for (c, ref), v in vals.get("F1", {}).items()}}},
        {"file": "02_Bank.xlsx", "sheets": ["Post"],
         "values": {"Post": {ref: v for (c, ref), v in vals.get("F2", {}).items()}}},
    ])
    shutil.move(tmp, os.path.join(OUT, "08_Dash.xlsx"))
    print("  -> 08_Dash.xlsx", os.path.getsize(os.path.join(OUT, "08_Dash.xlsx")) // 1024, "KB")

    # make sure the copied files point at this folder's 01_Entry (already do)
    print("done")


if __name__ == "__main__":
    main("--fast" in sys.argv)
