"""Put back what openpyxl drops when it saves a workbook: shapes, text boxes, connectors and grouped
drawings (copied from the original drawing XML, moved with inserted / deleted rows and columns) and
sparklines (the sheet's x14 extension). Runs on the saved file right after openpyxl wrote it."""
import os
import posixpath
import re
import zipfile

SPREADSHEET_DRAWING = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
DRAWING_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
DRAWING_CT = "application/vnd.openxmlformats-officedocument.drawing+xml"
SPARKLINES_URI = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"
# elements that come after <drawing> in a worksheet (it must be placed before the first of them)
AFTER_DRAWING = ("legacyDrawing", "legacyDrawingHF", "drawingHF", "picture", "oleObjects", "controls",
                 "webPublishItems", "tableParts", "extLst")


def _rels_path(part):
    d, b = posixpath.split(part)
    return posixpath.join(d, "_rels", b + ".rels")


def _resolve(part, target):
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(part), target))


def _rels(files, part):
    data = files.get(_rels_path(part))
    if not data:
        return {}
    out = {}
    for rel in re.findall(r"<Relationship\b[^>]*>", data.decode("utf-8")):
        rid, target, kind = (re.search(fr'{a}="([^"]*)"', rel) for a in ("Id", "Target", "Type"))
        if rid and target:
            out[rid.group(1)] = (target.group(1), kind.group(1) if kind else "")
    return out


def _drawing_of(files, sheet_part):
    """Path of the drawing part a sheet uses, or None."""
    xml = files[sheet_part].decode("utf-8")
    m = re.search(r"<(?:\w+:)?drawing\b[^>]*\br:id=\"([^\"]+)\"", xml)
    if not m:
        return None
    rel = _rels(files, sheet_part).get(m.group(1))
    return _resolve(sheet_part, rel[0]) if rel else None


def _prefix(xml, ns):
    m = re.search(r'xmlns:(\w+)="' + re.escape(ns) + '"', xml)
    return m.group(1) if m else None


def _ns_decls(xml):
    root = re.search(r"<(?:\w+:)?wsDr\b[^>]*>", xml)
    return dict(re.findall(r'xmlns:(\w+)="([^"]*)"', root.group(0))) if root else {}


def shape_anchors(drawing_xml):
    """Anchors holding shapes (not pictures / charts, which openpyxl keeps itself) -> (prefix, [xml])."""
    p = _prefix(drawing_xml, SPREADSHEET_DRAWING)
    if p is None:
        return None, []
    pat = re.compile(rf"<{p}:(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b.*?</{p}:\1>", re.S)
    out = []
    for m in pat.finditer(drawing_xml):
        a = m.group(0)
        if not re.search(rf"<{p}:(sp|grpSp|cxnSp)\b", a):
            continue
        if re.search(r'\br:(embed|link|id)="', a):    # needs another part (an image fill…): left out
            continue
        out.append(a)
    return p, out


def _remap(anchor, p, rows, cols):
    """Move an anchor's cells with inserted / deleted rows and columns; None if its cells were deleted."""
    def fix(tag, fn, text):
        def sub(m):
            new = fn(int(m.group(1)))
            if new is None:
                raise LookupError
            return f"<{p}:{tag}>{new}</{p}:{tag}>"
        return re.sub(rf"<{p}:{tag}>(\d+)</{p}:{tag}>", sub, text)
    try:
        if rows:
            anchor = fix("row", rows, anchor)
        if cols:
            anchor = fix("col", cols, anchor)
    except LookupError:
        return None
    return anchor


def _insert_drawing_ref(sheet_xml, rid):
    if not re.search(r"<(\w+:)?worksheet\b[^>]*xmlns:r=", sheet_xml):
        sheet_xml = re.sub(r"<((?:\w+:)?worksheet)\b", rf'<\1 xmlns:r="{REL_NS}"', sheet_xml, count=1)
    tag = f'<drawing r:id="{rid}"/>'
    for name in AFTER_DRAWING:
        m = re.search(rf"<(?:\w+:)?{name}\b", sheet_xml)
        if m:
            return sheet_xml[:m.start()] + tag + sheet_xml[m.start():]
    i = sheet_xml.rindex("</")
    return sheet_xml[:i] + tag + sheet_xml[i:]


def _add_rel(files, part, rel_type, target):
    path = _rels_path(part)
    data = files.get(path, b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships '
                           b'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>')
    xml = data.decode("utf-8")
    used = set(re.findall(r'Id="([^"]+)"', xml))
    n = 1
    while f"rId{n}" in used:
        n += 1
    rid = f"rId{n}"
    xml = xml.replace("</Relationships>", f'<Relationship Id="{rid}" Type="{rel_type}" Target="{target}"/></Relationships>')
    files[path] = xml.encode("utf-8")
    return rid


def _sparklines(sheet_xml):
    m = re.search(rf'<((?:\w+:)?ext)\b[^>]*uri="{re.escape(SPARKLINES_URI)}"[^>]*>.*?</\1>', sheet_xml, re.S)
    return m.group(0) if m else None


def restore(src, dst, names=None, maps=None):
    """src: the original file (path or its bytes). names: {sheet name in src: name in dst};
    maps: {dst sheet name: (row fn, col fn)} for rebuilt sheets
    (fn(old index) -> new index or None). Returns how many shapes / sparkline groups were put back."""
    from sheet_io import _sheet_files
    try:
        import io
        with zipfile.ZipFile(io.BytesIO(src) if isinstance(src, bytes) else src) as z:
            sfiles = {n: z.read(n) for n in z.namelist()}
        with zipfile.ZipFile(dst) as z:
            infos = z.infolist()
            dfiles = {i.filename: z.read(i.filename) for i in infos}
    except (OSError, zipfile.BadZipFile, KeyError):
        return 0
    if "xl/workbook.xml" not in sfiles:
        return 0
    s_where, d_where = _sheet_files(sfiles), _sheet_files(dfiles)
    names = names or {n: n for n in s_where}
    maps = maps or {}
    restored, order = 0, [i.filename for i in infos]
    for s_name, s_part in s_where.items():
        d_name = names.get(s_name)
        d_part = d_where.get(d_name) if d_name else None
        if not d_part or s_part not in sfiles or d_part not in dfiles:
            continue
        rows, cols = maps.get(d_name, (None, None))
        sheet_xml = dfiles[d_part].decode("utf-8")
        changed = False
        s_drawing = _drawing_of(sfiles, s_part)
        if s_drawing and s_drawing in sfiles:
            src_xml = sfiles[s_drawing].decode("utf-8")
            p, anchors = shape_anchors(src_xml)
            anchors = [a for a in (_remap(a, p, rows, cols) for a in anchors) if a]
            if anchors:
                d_drawing = _drawing_of(dfiles, d_part)
                if d_drawing and d_drawing in dfiles:
                    dxml = dfiles[d_drawing].decode("utf-8")
                    dp = _prefix(dxml, SPREADSHEET_DRAWING) or "xdr"
                    have = _ns_decls(dxml)
                    extra = "".join(f' xmlns:{k}="{v}"' for k, v in _ns_decls(src_xml).items()
                                    if k not in have and k != p)
                    if extra:
                        dxml = re.sub(rf"<{dp}:wsDr\b", lambda m: m.group(0) + extra, dxml, count=1)
                    if dp != p:
                        anchors = [re.sub(rf"(</?){p}:", rf"\g<1>{dp}:", a) for a in anchors]
                    dxml = dxml.replace(f"</{dp}:wsDr>", "".join(anchors) + f"</{dp}:wsDr>")
                    dfiles[d_drawing] = dxml.encode("utf-8")
                else:
                    n = 1
                    while f"xl/drawings/drawing{n}.xml" in dfiles:
                        n += 1
                    d_drawing = f"xl/drawings/drawing{n}.xml"
                    decls = "".join(f' xmlns:{k}="{v}"' for k, v in _ns_decls(src_xml).items())
                    dfiles[d_drawing] = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<{p}:wsDr{decls}>'
                                         + "".join(anchors) + f"</{p}:wsDr>").encode("utf-8")
                    order.append(d_drawing)
                    ct = dfiles["[Content_Types].xml"].decode("utf-8")
                    ct = ct.replace("</Types>", f'<Override PartName="/{d_drawing}" ContentType="{DRAWING_CT}"/></Types>')
                    dfiles["[Content_Types].xml"] = ct.encode("utf-8")
                    rid = _add_rel(dfiles, d_part, DRAWING_REL,
                                   posixpath.relpath(d_drawing, posixpath.dirname(d_part)))
                    if _rels_path(d_part) not in order:
                        order.append(_rels_path(d_part))
                    sheet_xml = _insert_drawing_ref(sheet_xml, rid)
                    changed = True
                restored += len(anchors)
        spark = _sparklines(sfiles[s_part].decode("utf-8"))
        if spark and rows is None and cols is None and not _sparklines(sheet_xml):
            src_root = re.search(r"<(?:\w+:)?worksheet\b[^>]*>", sfiles[s_part].decode("utf-8")).group(0)
            decls = dict(re.findall(r'xmlns:(\w+)="([^"]*)"', src_root))
            for pre in set(re.findall(r"<(\w+):", spark)):     # namespaces the extension uses
                if f"xmlns:{pre}=" not in spark and pre in decls:
                    spark = re.sub(r"^<(\S+)", lambda m, pre=pre: f'{m.group(0)} xmlns:{pre}="{decls[pre]}"', spark,
                                   count=1)
            m = re.search(r"</((?:\w+:)?extLst)>", sheet_xml)
            if m:
                sheet_xml = sheet_xml[:m.start()] + spark + sheet_xml[m.start():]
            else:
                i = sheet_xml.rindex("</")
                sheet_xml = sheet_xml[:i] + "<extLst>" + spark + "</extLst>" + sheet_xml[i:]
            changed = True
            restored += 1
        if changed:
            dfiles[d_part] = sheet_xml.encode("utf-8")
    if not restored:
        return 0
    tmp = dst + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for n in order:
            if n in dfiles:
                z.writestr(n, dfiles[n])
    os.replace(tmp, dst)
    return restored
