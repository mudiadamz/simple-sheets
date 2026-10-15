"""Read .xlsb (Excel binary workbook, BIFF12 records in a zip) with the standard library: sheet names
and visibility, cell values, formulas (xlsb_formula; the saved result is kept when one can't be read),
cell formats (fonts, fills, borders, alignment, number formats), column widths, row heights, hidden rows /
columns and merged cells. Writing is in xlsb_write.py."""
import datetime as dt
import re
import struct
import zipfile

from formatting import DEFAULT_FONT, DEFAULT_SIZE, EPOCH, theme_colors

# record types
BUNDLE_SH, ROW_HDR, SST_ITEM, FMT, XF, BEGIN_CELL_XFS, END_CELL_XFS = 0x9C, 0x00, 0x13, 0x2C, 0x2F, 0x269, 0x26A
FONT, FILL, BORDER, COL_INFO, MERGE_CELL, SHR_FMLA, EXTERN_SHEET = 0x2B, 0x2D, 0x2E, 0x3C, 0xB0, 0x1AA, 0x16A
CELLS = {0x01: "blank", 0x02: "rk", 0x03: "err", 0x04: "bool", 0x05: "real", 0x06: "st", 0x07: "isst",
         0x08: "fstr", 0x09: "fnum", 0x0A: "fbool", 0x0B: "ferr", 0x3E: "rstr"}
ERRORS = {0x00: "#NULL!", 0x07: "#DIV/0!", 0x0F: "#VALUE!", 0x17: "#REF!", 0x1D: "#NAME?", 0x24: "#NUM!", 0x2A: "#N/A"}
DATE_IDS = set(range(14, 23)) | {45, 46, 47}
BORDER_STYLES = {1: "thin", 2: "medium", 3: "dashed", 4: "dotted", 5: "thick", 6: "double", 7: "hair",
                 8: "mediumDashed", 9: "dashDot", 10: "mediumDashDot", 11: "dashDotDot", 12: "mediumDashDotDot",
                 13: "slantDashDot"}
H_ALIGN = {1: "left", 2: "center", 3: "right", 5: "justify", 6: "center", 7: "justify"}


def records(data):
    """Yield (type, payload) from BIFF12 bytes."""
    i, n = 0, len(data)
    while i < n:
        b = data[i]
        i += 1
        rtype = b & 0x7F
        if b & 0x80:
            rtype |= (data[i] & 0x7F) << 7
            i += 1
        size, shift = 0, 0
        for _ in range(4):
            b = data[i]
            i += 1
            size |= (b & 0x7F) << shift
            shift += 7
            if not b & 0x80:
                break
        yield rtype, data[i:i + size]
        i += size


def wide_string(buf, pos=0):
    """XLWideString -> (text, next position)."""
    (count,) = struct.unpack_from("<I", buf, pos)
    pos += 4
    if count == 0xFFFFFFFF:                       # XLNullableWideString: null
        return "", pos
    return buf[pos:pos + count * 2].decode("utf-16-le", "replace"), pos + count * 2


def rk(v):
    if v & 0x02:
        n = v >> 2
        if n & 0x20000000:
            n -= 0x40000000
        x = float(n)
    else:
        x = struct.unpack("<d", struct.pack("<Q", (v & 0xFFFFFFFC) << 32))[0]
    return x / 100 if v & 0x01 else x


def fmt_number(x):
    return str(int(x)) if float(x).is_integer() and abs(x) < 1e15 else f"{x:.15g}"


def as_date(x):
    try:
        d = EPOCH + dt.timedelta(days=float(x))
    except (OverflowError, ValueError):
        return fmt_number(x)
    return d.strftime("%Y-%m-%d") if d.time() == dt.time() else d.strftime("%Y-%m-%d %H:%M:%S")


def is_date_format(code):
    code = re.sub(r'"[^"]*"|\[[^\]]*\]|\\.', "", code or "").lower()
    return bool(re.search(r"[dmyhs]", code)) and not re.search(r"[0#?]", code)


def color(buf, pos, theme):
    """BrtColor (8 bytes) -> '#rrggbb' or None (automatic)."""
    from openpyxl.styles.colors import Color

    from formatting import color_hex
    flags, index, tint_, r, g, b, _a = struct.unpack_from("<BBhBBBB", buf, pos)
    kind = flags >> 1
    tint_ = tint_ / 32767.0
    if kind == 2:
        c = Color(rgb="FF%02X%02X%02X" % (r, g, b), tint=tint_)
    elif kind == 3:
        c = Color(theme=index, tint=tint_)
    elif kind == 1:
        c = Color(indexed=index, tint=tint_)
    else:
        return None
    return color_hex(c, theme)


def read_styles(data, theme=None):
    """styles.bin -> (dates, styles): per cell format index, is it a date / its style dict (or None)."""
    from openpyxl.styles.numbers import BUILTIN_FORMATS
    theme = theme or theme_colors(None)
    fmts, fonts, fills, borders, xfs, in_cell = {}, [], [], [], [], False
    for t, p in records(data):
        if t == FMT:
            (ifmt,) = struct.unpack_from("<H", p, 0)
            fmts[ifmt], _ = wide_string(p, 2)
        elif t == FONT and len(p) >= 21:
            height, grbit, bls, _sss, uls = struct.unpack_from("<HHHHB", p, 0)
            name, _ = wide_string(p, 21)
            f = {}
            if bls >= 700:
                f["bold"] = True
            if grbit & 0x02:
                f["italic"] = True
            if grbit & 0x08:
                f["strike"] = True
            if uls:
                f["underline"] = True
            fg = color(p, 12, theme)
            if fg and fg != "#000000":
                f["fg"] = fg
            if height and height / 20 != DEFAULT_SIZE:
                f["size"] = height / 20
            if name and name != DEFAULT_FONT:
                f["font"] = name
            fonts.append(f)
        elif t == FILL and len(p) >= 20:
            (fls,) = struct.unpack_from("<I", p, 0)
            bg = color(p, 4, theme) if fls == 1 else None
            fills.append({"bg": bg} if bg else {})
        elif t == BORDER and len(p) >= 51:
            b = {}
            for k, side in enumerate(("top", "bottom", "left", "right")):
                pos = 1 + k * 10
                if p[pos] in BORDER_STYLES:
                    b[side] = [BORDER_STYLES[p[pos]], color(p, pos + 2, theme) or "#000000"]
            borders.append({"border": b} if b else {})
        elif t == BEGIN_CELL_XFS:
            in_cell = True
        elif t == END_CELL_XFS:
            in_cell = False
        elif t == XF and in_cell:
            _parent, ifmt, ifont, ifill, iborder, trot, indent, flags = struct.unpack_from("<HHHHHBBI", p, 0)
            code = fmts.get(ifmt) or BUILTIN_FORMATS.get(ifmt)
            st = {}
            for lst, idx in ((fonts, ifont), (fills, ifill), (borders, iborder)):
                if idx < len(lst):
                    st.update(lst[idx])
            if code and code != "General":
                st["numfmt"] = code
            if flags & 0x7 in H_ALIGN:
                st["align"] = H_ALIGN[flags & 0x7]
            v = (flags >> 3) & 0x7
            if v in (0, 1, 3, 4):
                st["valign"] = "top" if v == 0 else "center"
            if flags & 0x40:
                st["wrap"] = True
            if indent:
                st["indent"] = indent
            if trot and trot != 255:
                st["rotate"] = 90 - trot if 90 < trot <= 180 else trot
            if not flags & 0x1000:
                st["unlocked"] = True
            xfs.append((ifmt in DATE_IDS or (ifmt in fmts and is_date_format(fmts[ifmt])), st or None))
    return [d for d, _ in xfs], [s for _, s in xfs]


def _formula_part(p, pos, base, sheets):
    """BrtFmla*: after the value, grbitFlags (2) and CellParsedFormula (cce, rgce, cb, rgcb)."""
    import xlsb_formula
    pos += 2
    (cce,) = struct.unpack_from("<I", p, pos)
    rgce = p[pos + 4:pos + 4 + cce]
    return rgce, xlsb_formula.decode(rgce, base, sheets)


def read_sheet(data, strings, dates, max_rows, styles=None, sheets=None):
    """-> (rows, extras) where extras holds styles, cache, widths, heights, hidden rows / cols, merges."""
    import xlsb_formula
    rows, r = [], -1
    ex = {"styles": {}, "cache": {}, "widths": {}, "heights": {}, "hidden_rows": set(), "hidden_cols": set(),
          "merges": [], "shared": []}
    pending = []                                  # cells waiting for a shared formula: (r, c, row of its first cell)
    styles = styles or []
    for t, p in records(data):
        if t == ROW_HDR:
            (r,) = struct.unpack_from("<I", p, 0)
            if len(p) >= 12 and r < max_rows:
                (height, flags) = struct.unpack_from("<HH", p, 8)
                if flags & 0x1000:
                    ex["hidden_rows"].add(r)
                if flags & 0x2000 and height:
                    ex["heights"][r] = height / 20
            continue
        if t == COL_INFO and len(p) >= 18:
            c1, c2, width, _ixfe, flags = struct.unpack_from("<IIIIH", p, 0)
            for c in range(c1, min(c2, 16383) + 1):
                if flags & 0x02 or width:
                    ex["widths"][c] = round(width / 256, 2)
                if flags & 0x01:
                    ex["hidden_cols"].add(c)
            continue
        if t == MERGE_CELL:
            r1, r2, c1, c2 = struct.unpack_from("<IIII", p, 0)
            ex["merges"].append((r1, c1, r2, c2))
            continue
        if t == SHR_FMLA:
            r1, r2, c1, c2 = struct.unpack_from("<IIII", p, 0)
            (cce,) = struct.unpack_from("<I", p, 16)
            ex["shared"].append(((r1, c1, r2, c2), p[20:20 + cce]))
            continue
        kind = CELLS.get(t)
        if kind is None or r < 0 or r >= max_rows:
            continue
        col, style = struct.unpack_from("<II", p, 0)
        style &= 0xFFFFFF
        if style < len(styles) and styles[style]:
            ex["styles"][(r, col)] = styles[style]
        v, end = "", 8
        if kind == "rk":
            v = rk(struct.unpack_from("<I", p, 8)[0])
        elif kind in ("real", "fnum"):
            v, end = struct.unpack_from("<d", p, 8)[0], 16
        elif kind in ("bool", "fbool"):
            v, end = ("TRUE" if p[8] else "FALSE"), 9
        elif kind in ("err", "ferr"):
            v, end = ERRORS.get(p[8], "#VALUE!"), 9
        elif kind in ("st", "fstr"):
            v, end = wide_string(p, 8)
        elif kind == "rstr":                      # rich text: the plain text is kept
            v, end = wide_string(p, 9)
        elif kind == "isst":
            (k,) = struct.unpack_from("<I", p, 8)
            v = strings[k] if k < len(strings) else ""
        if isinstance(v, float):
            v = as_date(v) if style < len(dates) and dates[style] else fmt_number(v)
        if kind.startswith("f"):
            try:
                rgce, text = _formula_part(p, end, (r, col), sheets)
            except struct.error:
                rgce, text = b"", None
            if text is not None:
                ex["cache"][(r, col)] = v
                v = "=" + text
            elif xlsb_formula.shared_ref(rgce) is not None:
                ex["cache"][(r, col)] = v
                pending.append((r, col, v))
        if v == "":
            continue
        while len(rows) <= r:
            rows.append([])
        row = rows[r]
        while len(row) <= col:
            row.append("")
        row[col] = v
    for r, c, _cached in pending:                 # parts of shared formulas: decoded relative to each cell
        for (r1, c1, r2, c2), rgce in ex["shared"]:
            if r1 <= r <= r2 and c1 <= c <= c2:
                text = xlsb_formula.decode(rgce, (r, c), sheets)
                if text is not None:
                    rows[r][c] = "=" + text
                break
    return rows, ex


def read_extern_sheets(workbook_bin, sheet_names):
    """XTI index -> sheet name (only references to this workbook's own sheets; others None)."""
    out = []
    for t, p in records(workbook_bin):
        if t == EXTERN_SHEET:
            (n,) = struct.unpack_from("<I", p, 0)
            for k in range(n):
                _book, first, last = struct.unpack_from("<iii", p, 4 + k * 12)
                ok = first == last and 0 <= first < len(sheet_names)
                out.append(sheet_names[first] if ok else None)
    return out


def load(path):
    import sheet_io as io
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        rels = z.read("xl/_rels/workbook.bin.rels").decode("utf-8", "replace")
        targets = dict(re.findall(r'Id="([^"]+)"[^>]*?Target="([^"]+)"', rels))
        targets.update({k: v for v, k in re.findall(r'Target="([^"]+)"[^>]*?Id="([^"]+)"', rels)})
        strings = []
        if "xl/sharedStrings.bin" in names:
            for t, p in records(z.read("xl/sharedStrings.bin")):
                if t == SST_ITEM:
                    strings.append(wide_string(p, 1)[0])
        theme = theme_colors(z.read("xl/theme/theme1.xml") if "xl/theme/theme1.xml" in names else None)
        dates, styles = read_styles(z.read("xl/styles.bin"), theme) if "xl/styles.bin" in names else ([], [])
        wb = z.read("xl/workbook.bin")
        bundles = []
        for t, p in records(wb):
            if t != BUNDLE_SH:
                continue
            state, _tab = struct.unpack_from("<II", p, 0)
            rel, pos = wide_string(p, 8)
            name, _ = wide_string(p, pos)
            bundles.append((state, rel, name))
        xti = read_extern_sheets(wb, [b[2] for b in bundles])
        out = {}
        for state, rel, name in bundles:
            target = targets.get(rel, "")
            part = "xl/" + target.lstrip("/").removeprefix("xl/")
            if part not in names:
                continue
            rows, ex = read_sheet(z.read(part), strings, dates, io.MAX_ROWS, styles, xti)
            sheet = io.Sheet(rows, ex["styles"], ex["widths"], ex["merges"], ex["heights"])
            sheet.cache = ex["cache"]
            sheet.hidden_rows, sheet.hidden_cols = ex["hidden_rows"], ex["hidden_cols"]
            sheet.hidden = state != 0
            sheet.reset_structure()
            out[name] = sheet
    return out


def unsupported(path):
    """What an .xlsb holds that Simple Sheets doesn't read or write (lost when saving over it)."""
    kinds = [("xl/drawings/", "pictures or charts"), ("xl/charts/", "charts"), ("xl/comments", "comments"),
             ("xl/pivot", "pivot tables"), ("xl/tables/", "tables"), ("xl/vbaProject", "macros"),
             ("xl/externalLinks/", "links to other workbooks"), ("xl/persons/", "comment threads"),
             ("xl/threadedComments/", "comment threads")]
    out = []
    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            for prefix, label in kinds:
                if any(n.startswith(prefix) for n in names) and label not in out:
                    out.append(label)
            if any(t == 0x27 for t, _ in records(z.read("xl/workbook.bin"))):
                out.append("named ranges")
            for n in names:
                if n.startswith("xl/worksheets/") and n.endswith(".bin"):
                    seen = {t for t, _ in records(z.read(n))}
                    for rtype, label in ((0x1CD, "conditional formatting"), (0x23D, "data validation"),
                                         (0xA1, "filters"), (0x1EE, "links")):
                        if rtype in seen and label not in out:
                            out.append(label)
    except (OSError, KeyError, zipfile.BadZipFile, IndexError):
        pass
    return out
