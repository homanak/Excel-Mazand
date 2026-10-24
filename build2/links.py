# -*- coding: utf-8 -*-
"""Creates the xl/externalLinks parts that xlsxwriter cannot write.

Also small helpers to harvest real cell values (so the delivered workbook shows
numbers before Excel refreshes the links).
"""
import json
import os
import re
import shutil
import zipfile
from xml.sax.saxutils import escape

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"


def _sheet_data_xml(values):
    rows = {}
    for ref, val in values.items():
        m = re.match(r"([A-Z]+)(\d+)$", ref)
        if not m:
            continue
        rows.setdefault(int(m.group(2)), []).append((ref, val))
    out = []
    for r in sorted(rows):
        cells = []
        for ref, val in sorted(rows[r], key=lambda t: (len(t[0]), t[0])):
            if val is None or val == "":
                continue
            if isinstance(val, bool):
                val = 1 if val else 0
            if isinstance(val, (int, float)):
                cells.append(f'<cell r="{ref}"><v>{val}</v></cell>')
            else:
                cells.append(f'<cell r="{ref}" t="str"><v>{escape(str(val))}</v></cell>')
        if cells:
            out.append(f'<row r="{r}">{"".join(cells)}</row>')
    return "".join(out)


def patch(path, links):
    """links = [{'file': '01_Entry.xlsx', 'sheets': ['Book', ...],
                 'values': {'Book': {'A10': '1405/07/01'}, ...}}, ...]"""
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zin:
        items = {n: zin.read(n) for n in zin.namelist()}

    wb_xml = items["xl/workbook.xml"].decode("utf-8")
    rels = items["xl/_rels/workbook.xml.rels"].decode("utf-8")
    ct = items["[Content_Types].xml"].decode("utf-8")
    if "externalLink+xml" in ct:
        return

    ids = [int(x) for x in re.findall(r'Id="rId(\d+)"', rels)]
    nid = (max(ids) + 1) if ids else 1
    refs = []
    for i, link in enumerate(links, start=1):
        rid = f"rId{nid}"
        nid += 1
        refs.append(f'<externalReference r:id="{rid}"/>')
        rels = rels.replace(
            "</Relationships>",
            f'<Relationship Id="{rid}" Type="{REL}/externalLink" '
            f'Target="externalLinks/externalLink{i}.xml"/></Relationships>')

        names = "".join(f'<sheetName val="{escape(s)}"/>' for s in link["sheets"])
        values = link.get("values") or {}
        data = []
        for idx, sname in enumerate(link["sheets"]):
            cells = _sheet_data_xml(values.get(sname) or {})
            data.append(f'<sheetData sheetId="{idx}">{cells}</sheetData>' if cells
                        else f'<sheetData sheetId="{idx}"/>')
        ext = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
               f'<externalLink xmlns="{NS}" xmlns:r="{REL}">'
               f'<externalBook r:id="rId1"><sheetNames>{names}</sheetNames>'
               f'<sheetDataSet>{"".join(data)}</sheetDataSet></externalBook></externalLink>')
        items[f"xl/externalLinks/externalLink{i}.xml"] = ext.encode("utf-8")
        items[f"xl/externalLinks/_rels/externalLink{i}.xml.rels"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{REL}/externalLinkPath" '
            f'Target="{escape(link["file"])}" TargetMode="External"/></Relationships>'
        ).encode("utf-8")
        ct = ct.replace("</Types>",
                        f'<Override PartName="/xl/externalLinks/externalLink{i}.xml" '
                        f'ContentType="{CT}"/></Types>')

    block = "<externalReferences>" + "".join(refs) + "</externalReferences>"
    if "<externalReferences>" not in wb_xml:
        wb_xml = wb_xml.replace("</sheets>", "</sheets>" + block, 1)

    items["xl/workbook.xml"] = wb_xml.encode("utf-8")
    items["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")
    items["[Content_Types].xml"] = ct.encode("utf-8")

    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in items.items():
            zout.writestr(name, data)
    shutil.move(tmp, path)


# ------------------------------------------------------------------ harvest
def _plain(v):
    try:
        arr = getattr(v, "value", v)
        while hasattr(arr, "shape") and getattr(arr, "shape", None):
            arr = arr[0, 0] if arr.size else None
        if arr is None:
            return 0
        if isinstance(arr, (int, float)):
            f = float(arr)
            return int(f) if f.is_integer() else f
        s = str(arr)
        if s in ("empty", "", "None"):
            return ""
        return s
    except Exception:
        return 0


def harvest(workbook_path, sheets_cells, out_json):
    """sheets_cells = {'Book': ['A10', ...], 'Post': [...]}"""
    import formulas
    model = formulas.ExcelModel().loads(workbook_path).finish()
    sol = model.calculate()
    name = os.path.basename(workbook_path)
    out = {}
    miss = 0
    for sheet, cells in sheets_cells.items():
        out[sheet] = {}
        for ref in cells:
            key = f"'[{name}]{sheet.upper()}'!{ref}"
            if key in sol:
                out[sheet][ref] = _plain(sol[key])
            else:
                miss += 1
    with open(out_json, "w", encoding="utf-8") as fh:
        json.dump(out, fh, ensure_ascii=False)
    return out, miss


def harvest_cells(sheet, cols, first, last):
    return [f"{c}{r}" for r in range(first, last + 1) for c in cols]
