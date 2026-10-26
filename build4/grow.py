# -*- coding: utf-8 -*-
"""Row engine — expands every entry log to its final capacity and rebuilds the
journal (Book) accordingly.

The v03 series was sized for 300 weighbridge rows / 150 basket rows / 150
payment rows → a 1,799-line journal.  The new series is sized for

    Scale   10 …  10009     10,000 weighbridge rows
    Crate   10 …   2509      2,500 basket rows
    Pay     10 …   2509      2,500 payment rows
    Book    10 …  50009     50,000 journal lines   (what the client asked for)

and the five report files + dashboard mirror all 50,000 lines.

Every block is grown by *cloning its last template row* and re-pointing the
input row number, so no formula is invented.
"""
import re

from openpyxl.utils import get_column_letter

# ------------------------------------------------------------------ capacity
SCALE_LAST = 10009          # Scale!10 .. Scale!10009   → 10,000 rows
PAY_LAST = 2509             # Pay!10   .. Pay!2509     →  2,500 rows
CRATE_LAST = 2509           # Crate!10 .. Crate!2509   →  2,500 rows
DATA_LAST = 50009           # mirror sheets (03/04/05/06 Data)
BOOK_LAST = 50009           # Book!10  .. Book!50009   → 50,000 rows

def book_blocks():
    """(source sheet, input rows, lines per row, template row) — read lazily so
    a test run can shrink the capacities."""
    return [
        ("Scale", SCALE_LAST - 9, 4, 1206),   # weighbridge: main / counter / POS / counter
        ("Pay", PAY_LAST - 9, 2, 1508),       # payment: first line / second line
        ("Crate", CRATE_LAST - 9, 2, 1808),   # basket: first line / second line
    ]

_REF = re.compile(r"\$([A-Z]{1,3})\$(\d+)")
_ROWREF = re.compile(r"(\$?[A-Z]{1,3}\$?)10\b")


# ---------------------------------------------------------------- primitives
def clone_rows(ws, src_row, first, last, cols=None):
    """Copy ``src_row`` (value + style + height) down to ``first..last``."""
    cols = cols or range(1, ws.max_column + 1)
    tpl = {c: (ws.cell(src_row, c).value, ws.cell(src_row, c)._style) for c in cols}
    height = ws.row_dimensions[src_row].height
    for r in range(first, last + 1):
        for c, (val, sty) in tpl.items():
            cell = ws.cell(r, c)
            # a cloned row must point at *itself*, not at the template row
            cell.value = _shift_row(val, src_row, r) if isinstance(val, str) else val
            cell._style = sty
        ws.row_dimensions[r].height = height


def _shift_row(text, old, new):
    """Re-point row-relative references (``$A10``) to another row."""
    if not isinstance(text, str) or not text.startswith("="):
        return text
    return re.sub(r"(\$?[A-Z]{1,3}\$?)" + str(old) + r"\b",
                  lambda m: m.group(1) + str(new), text)


def rebuild_book(bk):
    """Rewrite every journal block with clones of its own template lines."""
    # 1) read every template (value + style) *before* touching the sheet
    plan = []
    for sheet, count, per, tpl_row in book_blocks():
        tpl = [(bk.cell(tpl_row + i, c).value, bk.cell(tpl_row + i, c)._style,
                bk.row_dimensions[tpl_row + i].height)
               for i in range(per) for c in range(1, 18)]
        plan.append((sheet, count, per, tpl))
    total_style = [(bk.cell(1810, c)._style) for c in range(1, 18)]
    total_h = bk.row_dimensions[1810].height

    # 2) write the blocks one after another
    r = 10
    for sheet, count, per, tpl in plan:
        for i in range(count):
            src = 10 + i
            for k in range(per):
                for c in range(1, 18):
                    val, sty, h = tpl[k * 17 + (c - 1)]
                    cell = bk.cell(r, c)
                    cell.value = _shift_row(val, 309 if sheet == "Scale" else 159, src)
                    cell._style = sty
                bk.row_dimensions[r].height = h
                r += 1
    assert r - 1 == BOOK_LAST, (r, BOOK_LAST)
    return total_style, total_h


def _style_row(ws, src, dst, cols):
    for c in cols:
        ws.cell(dst, c)._style = ws.cell(src, c)._style
    ws.row_dimensions[dst].height = ws.row_dimensions[src].height


def grow_entry(wb):
    """Scale / Crate / Pay → capacity, then rebuild Book + its grand total."""
    scale = wb["Scale"]
    clone_rows(scale, 309, 310, SCALE_LAST)

    for name in ("Crate", "Pay"):
        ws = wb[name]
        clone_rows(ws, 159, 160, PAY_LAST)

    bk = wb["Book"]
    for rng in ("A1810:J1810", "N1810:O1810"):
        try:
            bk.unmerge_cells(rng)
        except ValueError:
            pass
    total_style, total_h = rebuild_book(bk)

    # grand total of the journal, just under the last data row
    tot = BOOK_LAST + 1
    for c in range(1, 18):
        bk.cell(tot, c)._style = total_style[c - 1]
    bk.row_dimensions[tot].height = total_h
    bk.merge_cells(f"A{tot}:J{tot}")
    bk.cell(tot, 1).value = "جمع کل دفتر  (باید برابر باشد)"
    for c, letter in ((11, "K"), (12, "L"), (13, "M")):
        bk.cell(tot, c).value = f"=SUM({letter}10:{letter}{BOOK_LAST})"
    bk.cell(tot, 16).value = f'=IF(ROUND($L${tot}-$M${tot},0)=0,"✓","⚠")'

    bk.auto_filter.ref = f"A8:Q{BOOK_LAST}"
    bk.print_area = f"A1:Q{tot}"
    return SCALE_LAST, PAY_LAST, BOOK_LAST


# --------------------------------------------------------- mirror / report
def grow_data(ws, new_last, first=10):
    """Grow a mirror sheet (03/04/05/06 ``Data``) from its row-10 template."""
    cols = range(1, ws.max_column + 1)
    tpl = {c: (ws.cell(first, c).value, ws.cell(first, c)._style) for c in cols}
    height = ws.row_dimensions[first].height
    for r in range(first, new_last + 1):
        for c, (val, sty) in tpl.items():
            cell = ws.cell(r, c)
            if isinstance(val, str) and val.startswith("="):
                v = re.sub(r"(\$?[A-Z]{1,3}\$?)" + str(first) + r"\b",
                           lambda m: m.group(1) + str(r), val)
                v = v.replace("$9)", f"${r - 1})")
                val = v
            cell.value = val
            cell._style = sty
        ws.row_dimensions[r].height = height
    return new_last


_RANGE = re.compile(r"(\$?[A-Z]{1,3}\$?10:\$?[A-Z]{1,3}\$?)"
                    r"(1807|1808|1809|157|158|159|307|308|309)\b")


def reseat_ranges(wb, mapping):
    """Rewrite *whole-log* ranges (start row 10) to the new capacity."""
    n = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if not isinstance(v, str) or not v.startswith("="):
                    continue
                new = _RANGE.sub(lambda m: m.group(1) + str(mapping[int(m.group(2))])
                                 if int(m.group(2)) in mapping else m.group(0), v)
                if new != v:
                    cell.value = new
                    n += 1
    return n


ENTRY_MAP = {1807: BOOK_LAST, 1808: BOOK_LAST, 1809: BOOK_LAST,
             157: PAY_LAST, 158: PAY_LAST, 159: PAY_LAST,
             307: SCALE_LAST, 308: SCALE_LAST, 309: SCALE_LAST}
