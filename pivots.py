"""Real Excel pivot tables. Data → Pivot table makes a summary sheet (the values are filled in here) and,
when saved, the pivot-table parts Excel needs: a pivot cache pointing at the source range and a pivot table
definition (rows, optional columns, one value field). Excel rebuilds the pivot when it opens the file
(refreshOnLoad), so it stays a live, refreshable Excel pivot table. Data → Refresh pivot tables
recalculates the summaries here.

App.pivots: [{"sheet": summary sheet, "source": (sheet, r1, c1, r2, c2), "rows": col index, "cols": col index
or None, "values": col index, "agg": "Sum"|"Count"|"Average"|"Min"|"Max", "saved": bool}]"""
import os
import posixpath
import re
import zipfile
from xml.sax.saxutils import quoteattr

import keep_parts
from formula import col_letters, quote_sheet

NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' \
     'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
CACHE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition"
TABLE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable"
CACHE_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"
TABLE_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"
SUBTOTAL = {"Sum": None, "Count": "count", "Average": "average", "Min": "min", "Max": "max"}


def source_rows(app, spec):
    """Header + data rows of the source range, as shown."""
    name, r1, c1, r2, c2 = spec["source"]
    s = app.sheets[name]
    eng = app.engine
    out = []
    for r in range(r1, min(r2, len(s.rows) - 1) + 1):
        row = []
        for c in range(c1, c2 + 1):
            t = s.rows[r][c] if c < len(s.rows[r]) else ""
            row.append(eng.display(s, r, c) if t.startswith("=") and len(t) > 1 else t)
        out.append(row)
    return out


def table(app, spec):
    """The summary values (same layout the app shows)."""
    import data_more
    rows = source_rows(app, spec)
    header = rows[0]
    col = header[spec["cols"]] if spec.get("cols") is not None else None
    return data_more.make_pivot(rows[1:], header, header[spec["rows"]], header[spec["values"]], spec["agg"], col)


def refresh(app):
    """Data → Refresh pivot tables: recalculate every summary sheet from its source (one undo step)."""
    changes = []
    for spec in getattr(app, "pivots", []):
        s = app.sheets.get(spec["sheet"])
        if s is None or spec["source"][0] not in app.sheets:
            continue
        new = table(app, spec)
        h = max(len(s.rows), len(new))
        w = max(max((len(r) for r in s.rows), default=0), max(len(r) for r in new))
        for r in range(h):
            for c in range(w):
                v = new[r][c] if r < len(new) and c < len(new[r]) else ""
                old = s.rows[r][c] if r < len(s.rows) and c < len(s.rows[r]) else ""
                if v != old:
                    changes.append(("v", s, r, c, v))
    if changes:
        app.apply_cells(changes)
    return len(changes)


def _items(values):
    """Distinct values in first-seen order, numbers before text like Excel's shared items."""
    seen = []
    for v in values:
        if v not in seen:
            seen.append(v)
    return seen


def _shared(values):
    from formatting import parse_number
    nums = [parse_number(v) for v in values if v != ""]
    if values and all(n is not None for n in nums) and nums:
        return (f'<sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" '
                f'minValue="{min(nums):g}" maxValue="{max(nums):g}"/>')
    items = _items([v for v in values])
    body = "".join("<m/>" if v == "" else f"<s v={quoteattr(v)}/>" for v in items)
    blank = ' containsBlank="1"' if "" in items else ""
    return f'<sharedItems{blank} count="{len(items)}">{body}</sharedItems>'


def cache_xml(app, spec):
    rows = source_rows(app, spec)
    name, r1, c1, r2, c2 = spec["source"]
    ref = f"{col_letters(c1)}{r1 + 1}:{col_letters(c2)}{r1 + len(rows)}"
    fields = []
    for j, head in enumerate(rows[0]):
        col = [r[j] if j < len(r) else "" for r in rows[1:]]
        fields.append(f'<cacheField name={quoteattr(head or f"Column{j + 1}")} numFmtId="0">{_shared(col)}</cacheField>')
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheDefinition {NS} saveData="0" '
            'refreshOnLoad="1" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" recordCount="0">'
            f'<cacheSource type="worksheet"><worksheetSource ref="{ref}" sheet={quoteattr(name)}/></cacheSource>'
            f'<cacheFields count="{len(fields)}">{"".join(fields)}</cacheFields></pivotCacheDefinition>')


def table_xml(app, spec, cache_id, number):
    rows = source_rows(app, spec)
    header, data = rows[0], rows[1:]
    n = len(header)
    rkeys = _items([r[spec["rows"]] for r in data])
    ckeys = _items([r[spec["cols"]] for r in data]) if spec.get("cols") is not None else None
    fields = []
    for j in range(n):
        if j == spec["rows"]:
            col_vals = _items([r[j] for r in data])
            items = "".join(f'<item x="{k}"/>' for k in range(len(col_vals))) + '<item t="default"/>'
            fields.append(f'<pivotField axis="axisRow" showAll="0"><items count="{len(col_vals) + 1}">{items}</items>'
                          "</pivotField>")
        elif ckeys is not None and j == spec["cols"]:
            items = "".join(f'<item x="{k}"/>' for k in range(len(ckeys))) + '<item t="default"/>'
            fields.append(f'<pivotField axis="axisCol" showAll="0"><items count="{len(ckeys) + 1}">{items}</items>'
                          "</pivotField>")
        elif j == spec["values"]:
            fields.append('<pivotField dataField="1" showAll="0"/>')
        else:
            fields.append('<pivotField showAll="0"/>')
    row_items = "".join("<i><x/></i>" if k == 0 else f'<i><x v="{k}"/></i>' for k in range(len(rkeys)))
    row_items += '<i t="grand"><x/></i>'
    if ckeys is not None:
        col_items = "".join("<i><x/></i>" if k == 0 else f'<i><x v="{k}"/></i>' for k in range(len(ckeys)))
        col_items += '<i t="grand"><x/></i>'
        cols = (f'<colFields count="1"><field x="{spec["cols"]}"/></colFields>'
                f'<colItems count="{len(ckeys) + 1}">{col_items}</colItems>')
        width, first_data_row, height = len(ckeys) + 2, 2, len(rkeys) + 3
    else:
        cols = '<colItems count="1"><i/></colItems>'
        width, first_data_row, height = 2, 1, len(rkeys) + 2
    sub = SUBTOTAL[spec["agg"]]
    sub_attr = f' subtotal="{sub}"' if sub else ""
    title = f'{spec["agg"]} of {header[spec["values"]]}'
    ref = f"A1:{col_letters(width - 1)}{height}"
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotTableDefinition {NS} '
            f'name="PivotTable{number}" cacheId="{cache_id}" applyNumberFormats="0" applyBorderFormats="0" '
            'applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" '
            'dataCaption="Values" updatedVersion="6" minRefreshableVersion="3" useAutoFormatting="1" '
            'itemPrintTitles="1" createdVersion="6" indent="0" outline="1" outlineData="1" multipleFieldFilters="0">'
            f'<location ref="{ref}" firstHeaderRow="1" firstDataRow="{first_data_row}" firstDataCol="1"/>'
            f'<pivotFields count="{n}">{"".join(fields)}</pivotFields>'
            f'<rowFields count="1"><field x="{spec["rows"]}"/></rowFields>'
            f'<rowItems count="{len(rkeys) + 1}">{row_items}</rowItems>{cols}'
            f'<dataFields count="1"><dataField name={quoteattr(title)} fld="{spec["values"]}" baseField="0" '
            f'baseItem="0"{sub_attr}/></dataFields>'
            '<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" '
            'showRowStripes="0" showColStripes="0" showLastColumn="1"/></pivotTableDefinition>')


def write(path, app):
    """Add the pivot parts for every pivot made here that isn't in the file yet. Returns how many."""
    from sheet_io import _sheet_files
    todo = [p for p in getattr(app, "pivots", []) if not p.get("saved") and p["sheet"] in app.sheets
            and p["source"][0] in app.sheets]
    if not todo:
        return 0
    with zipfile.ZipFile(path) as z:
        infos = z.infolist()
        files = {i.filename: z.read(i.filename) for i in infos}
    order = [i.filename for i in infos]
    where = _sheet_files(files)
    wb = files["xl/workbook.xml"].decode("utf-8")
    ids = [int(x) for x in re.findall(r'<pivotCache\b[^>]*cacheId="(\d+)"', wb)]
    ct = files["[Content_Types].xml"].decode("utf-8")
    entries = []
    done = 0
    for spec in todo:
        part = where.get(spec["sheet"])
        if not part:
            continue
        cache_id = max(ids + [0]) + 1
        ids.append(cache_id)
        n = 1
        while f"xl/pivotCache/pivotCacheDefinition{n}.xml" in files:
            n += 1
        cpart = f"xl/pivotCache/pivotCacheDefinition{n}.xml"
        k = 1
        while f"xl/pivotTables/pivotTable{k}.xml" in files:
            k += 1
        tpart = f"xl/pivotTables/pivotTable{k}.xml"
        files[cpart] = cache_xml(app, spec).encode("utf-8")
        files[tpart] = table_xml(app, spec, cache_id, k).encode("utf-8")
        order += [cpart, tpart]
        ct = ct.replace("</Types>", f'<Override PartName="/{cpart}" ContentType="{CACHE_CT}"/>'
                                    f'<Override PartName="/{tpart}" ContentType="{TABLE_CT}"/></Types>')
        rid = keep_parts._add_rel(files, "xl/workbook.xml", CACHE_REL, posixpath.relpath(cpart, "xl"))
        keep_parts._add_rel(files, tpart, CACHE_REL, posixpath.relpath(cpart, posixpath.dirname(tpart)))
        keep_parts._add_rel(files, part, TABLE_REL, posixpath.relpath(tpart, posixpath.dirname(part)))
        for p in (keep_parts._rels_path(tpart), keep_parts._rels_path(part)):
            if p not in order:
                order.append(p)
        entries.append(f'<pivotCache cacheId="{cache_id}" r:id="{rid}"/>')
        spec["saved"] = True
        done += 1
    if entries:
        if "<pivotCaches>" in wb:
            wb = wb.replace("</pivotCaches>", "".join(entries) + "</pivotCaches>")
        else:
            block = "<pivotCaches>" + "".join(entries) + "</pivotCaches>"
            m = re.search(r"<(smartTagPr|smartTagTypes|webPublishing|fileRecoveryPr|webPublishObjects|extLst)\b", wb)
            i = m.start() if m else wb.rindex("</workbook>")
            wb = wb[:i] + block + wb[i:]
        if 'xmlns:r=' not in wb.split(">", 2)[1]:
            wb = re.sub(r"<workbook\b", '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/'
                                        'relationships"', wb, count=1)
        files["xl/workbook.xml"] = wb.encode("utf-8")
    files["[Content_Types].xml"] = ct.encode("utf-8")
    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for name in order:
            if name in files:
                z.writestr(name, files[name])
    os.replace(tmp, path)
    return done


def ref_text(spec):
    name, r1, c1, r2, c2 = spec["source"]
    return f"{quote_sheet(name)}!{col_letters(c1)}{r1 + 1}:{col_letters(c2)}{r2 + 1}"
