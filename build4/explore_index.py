# -*- coding: utf-8 -*-
"""Index the nine workbooks so the browser can walk them like Excel.

    .venv/bin/python build4/explore_index.py [folder] [-o build4/_explore]

Writes one compact JSON cache that ``explore_server.py`` serves.  Nothing in the
workbooks is modified.

What is stored per sheet
  * a window of cells (row 1 … 220, column 1 … 32) with the *display* text,
    a style index and a link index — enough to draw the sheet faithfully
  * the formulas of those cells (kept apart, they are long)
  * row heights, column widths, merges, hidden rows/columns, print area
What is stored per file
  * every hyperlink (target file or target sheet + cell)
  * every named range and where it points (``OpTbl`` → ``Sync!$A$10:$B$33``)
  * the cross-sheet references found in the formulas (the data-flow edges)
"""
import gzip
import json
import os
import re
import sys
import time

import openpyxl
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SRC = os.path.join(ROOT, "Mazandasht_v04")
OUT = os.path.join(HERE, "_explore")

FILES = ["00_Master.xlsx", "01_Entry.xlsx", "02_Bank.xlsx", "03_Report.xlsx",
         "04_Inv.xlsx", "05_Bal.xlsx", "06_Buyers.xlsx", "07_Demo.xlsx",
         "08_Dash.xlsx"]
LABEL = {"00_Master.xlsx": "فایل مادر", "01_Entry.xlsx": "دفتر ثبت", "02_Bank.xlsx": "بانک",
         "03_Report.xlsx": "گزارش‌ها", "04_Inv.xlsx": "فاکتورها", "05_Bal.xlsx": "تراز",
         "06_Buyers.xlsx": "خریداران", "07_Demo.xlsx": "نمونه", "08_Dash.xlsx": "داشبورد"}

ROWS_FULL = 240          # sheets up to this many rows are stored whole
ROWS_BIG = 160           # bigger sheets: first N rows only
COLS = 34

# a sheet-qualified reference: Scale!$X10 · 'Bill'!$O$2 · [f.xlsx]Sync!A1
REF_RE = re.compile(r"(?:\[([^\]]+\.xlsx)\])?(?:'([^'!]+)'|([A-Za-z_\u0600-\u06FF][\w\u0600-\u06FF]*))!")

STYLE_KEYS = ("bg", "fg", "b", "sz", "al", "wrap", "bd")


def _color(c):
    try:
        rgb = c.rgb
    except Exception:
        return None
    if isinstance(rgb, str) and len(rgb) == 8 and rgb[2:] != "000000":
        return "#" + rgb[2:]
    return None


def _display(v):
    if isinstance(v, bool):
        return "✓" if v else "✗"
    if isinstance(v, float):
        if abs(v) >= 1000:
            return f"{v:,.0f}" if abs(v - round(v)) < 1e-9 else f"{v:,.2f}"
        if v.is_integer():
            return str(int(v))
        return f"{v:.2f}"
    if isinstance(v, int):
        return f"{v:,}"
    return str(v)


def _refs(formula, own_sheet):
    """Sheet names (and defined names) a formula points at."""
    out = []
    if not isinstance(formula, str) or not formula.startswith("="):
        return out
    for m in REF_RE.finditer(formula):
        sheet = m.group(2) or m.group(3)
        if not sheet or sheet.startswith(("SUM", "IF", "ROUND", "MAX", "MIN", "COUNT")):
            continue
        out.append((m.group(1), sheet))
    return out


def build(folder):
    t0 = time.time()
    files = []
    sheets_by_file = {}
    links_all = []
    named_all = []

    for fname in FILES:
        path = os.path.join(folder, fname)
        if not os.path.exists(path):
            print("  ! missing", fname)
            continue
        print(f"  {fname} …", end="", flush=True)
        wb = openpyxl.load_workbook(path)
        styles, style_ix = [], {}
        rtl_book = {}

        def style_of(cell):
            f = cell.font
            a = cell.alignment
            bg = _color(cell.fill) if cell.fill is not None and cell.fill.fill_type == "solid" else None
            fg = _color(f.color) if f.color is not None else None
            bd = bool(cell.border and cell.border.left and cell.border.left.style)
            al = a.horizontal or ("right" if rtl_book[ws.title] else "left")
            key = (bg, fg, bool(f.b), round(float(f.sz or 11), 1), al, bool(a.wrap_text), bd)
            got = style_ix.get(key)
            if got is None:
                got = len(styles)
                style_ix[key] = got
                styles.append([bg, fg, 1 if f.b else 0, round(float(f.sz or 11), 1),
                               al, 1 if a.wrap_text else 0, 1 if bd else 0])
            return got

        named = {}
        try:
            for name in wb.defined_names:
                dn = wb.defined_names[name]
                named[name] = dn.value if isinstance(dn.value, str) else str(dn.value)
        except Exception:
            pass
        named_all.append({"file": fname, "names": named})

        fobj = {"name": fname, "label": LABEL.get(fname, fname), "size": os.path.getsize(path),
                "sheets": [], "links": [], "named": named}

        for ws in wb.worksheets:
            rtl_book[ws.title] = bool(ws.sheet_view.rightToLeft)
            big = ws.max_row > ROWS_FULL
            cap_r = ROWS_BIG if big else min(ws.max_row, ROWS_FULL)
            cap_c = min(ws.max_column or 1, COLS)
            cells, formulas, links = {}, {}, []
            for row in ws.iter_rows(min_row=1, max_row=cap_r, max_col=cap_c):
                for cell in row:
                    v = cell.value
                    if v is None:
                        continue
                    ref = cell.coordinate
                    h = cell.hyperlink
                    li = None
                    if h is not None:
                        tgt = h.target or ""
                        loc = h.location or ""
                        kind = "file" if tgt.lower().endswith(".xlsx") else "sheet"
                        if not tgt and not loc:
                            kind = "none"
                        li = len(links)
                        links.append({"cell": ref, "target": tgt, "loc": loc, "kind": kind,
                                      "text": _display(v)[:90]})
                    if isinstance(v, str) and v.startswith("="):
                        formulas[ref] = v
                        txt = "ƒ"
                    else:
                        txt = _display(v)[:300]
                    cells[ref] = [txt, style_of(cell), li]
            widths = {}
            for col in range(1, cap_c + 1):
                letter = get_column_letter(col)
                d = ws.column_dimensions.get(letter)
                widths[letter] = round(d.width, 2) if d is not None and d.width else None
            heights = {}
            for r in range(1, cap_r + 1):
                d = ws.row_dimensions.get(r)
                if d is not None and d.height:
                    heights[str(r)] = round(d.height, 1)
            hidden_rows = [r for r in range(1, cap_r + 1)
                           if ws.row_dimensions.get(r) is not None and ws.row_dimensions[r].hidden]
            hidden_cols = [get_column_letter(c) for c in range(1, cap_c + 1)
                           if ws.column_dimensions.get(get_column_letter(c)) is not None
                           and ws.column_dimensions[get_column_letter(c)].hidden]

            # ---- dependencies of this sheet: local sheets + named ranges
            deps, dep_cells = {}, []
            for ref, formula in formulas.items():
                seen = set()
                for _, sheet in _refs(formula, ws.title):
                    if sheet == ws.title or sheet in seen:
                        continue
                    seen.add(sheet)
                    if sheet in wb.sheetnames:
                        deps[sheet] = deps.get(sheet, 0) + 1
                        if len(dep_cells) < 40:
                            dep_cells.append({"cell": f"{ws.title}!{ref}", "formula": formula,
                                              "to": sheet})
                for name in named:
                    if re.search(rf"\b{re.escape(name)}\b", formula) and name not in seen:
                        target = named[name] or ""
                        ts = target.split("!")[0].strip("'")
                        if ts and ts != ws.title:
                            seen.add(name)
                            deps[ts] = deps.get(ts, 0) + 1
                            if len(dep_cells) < 40:
                                dep_cells.append({"cell": f"{ws.title}!{ref}", "formula": formula,
                                                  "to": ts, "name": name})
            sheets_by_file.setdefault(fname, {})[ws.title] = {
                "rows": ws.max_row, "cols": ws.max_column, "big": big,
                "cap_rows": cap_r, "cap_cols": cap_c,
                "print": str(ws.print_area or ""), "rtl": bool(ws.sheet_view.rightToLeft),
                "state": ws.sheet_state, "freeze": str(ws.freeze_panes or ""),
                "merges": [str(m) for m in ws.merged_cells.ranges][:400],
                "widths": widths, "heights": heights,
                "hidden_rows": hidden_rows[:200], "hidden_cols": hidden_cols,
                "cells": cells, "formulas": formulas,
                "deps": deps, "dep_cells": dep_cells,
            }
            fobj["sheets"].append({"name": ws.title, "rows": ws.max_row, "cols": ws.max_column,
                                   "big": big, "state": ws.sheet_state,
                                   "links": len(links), "print": str(ws.print_area or ""),
                                   "formulas": len(formulas)})
            fobj["links"].extend([dict(x, sheet=ws.title) for x in links])
        fobj["styles"] = styles
        files.append(fobj)
        links_all.extend([dict(x, file=fname) for x in fobj["links"]])
        print(f" {len(fobj['sheets'])} sheets, {len(fobj['links'])} links, "
              f"{time.time() - t0:.0f}s")

    # ---------------------------------------------------------------- edges
    edges = []
    for f in files:
        for s in f["sheets"]:
            sh = sheets_by_file[f["name"]][s["name"]]
            for target, count in sh["deps"].items():
                edges.append({"from": [f["name"], s["name"]], "to": [f["name"], target],
                              "kind": "formula", "count": count,
                              "samples": [d for d in sh["dep_cells"] if d["to"] == target][:6]})
        for ln in f["links"]:
            tgt_file = f["name"]
            tgt_sheet = None
            if ln["kind"] == "file":
                tgt_file = ln["target"].split("#")[0].strip()
                loc = ln["loc"] or (ln["target"].split("#")[1] if "#" in ln["target"] else "")
                tgt_sheet = loc.split("!")[0].strip("'") if "!" in loc else None
            elif ln["kind"] == "sheet":
                tgt_sheet = (ln["loc"] or "").split("!")[0].strip("'") or None
            edges.append({"from": [f["name"], ln.get("sheet")], "to": [tgt_file, tgt_sheet],
                          "kind": "file" if ln["kind"] == "file" else "sheet", "count": 1,
                          "samples": [{"cell": f"{ln.get('sheet')}!{ln['cell']}",
                                       "text": ln["text"]}]})

    # the Sync bus: every file carries a locked copy of the master data
    for f in files:
        if f["name"] == "00_Master.xlsx":
            continue
        if "Sync" in sheets_by_file.get(f["name"], {}):
            edges.append({"from": ["00_Master.xlsx", "Sync"], "to": [f["name"], "Sync"],
                          "kind": "sync", "count": 1,
                          "samples": [{"cell": "00_Master.xlsx → Sync", "text": "کپی قفل‌شدهٔ دادهٔ مادر"}]})

    data = {"generated": time.strftime("%Y-%m-%d %H:%M"), "source": folder,
            "files": files, "sheets": sheets_by_file, "edges": edges,
            "named": named_all}
    return data


def main(argv):
    folder = argv[1] if len(argv) > 1 and not argv[1].startswith("-") else SRC
    out = OUT
    if "-o" in argv:
        out = argv[argv.index("-o") + 1]
    os.makedirs(out, exist_ok=True)
    data = build(folder)
    raw = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    path = os.path.join(out, "index.json")
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(raw)
    with gzip.open(path + ".gz", "wb") as fh:
        fh.write(raw.encode("utf-8"))
    nodes = sum(len(f["sheets"]) for f in data["files"])
    print(f"\nindex → {path}  ({len(raw) / 1e6:.1f} MB, gz {os.path.getsize(path + '.gz') / 1e6:.1f} MB)")
    print(f"files {len(data['files'])} · sheets {nodes} · links "
          f"{sum(len(f['links']) for f in data['files'])} · edges {len(data['edges'])}")


if __name__ == "__main__":
    main(sys.argv)
