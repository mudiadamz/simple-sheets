"""Save as .xlsb (Excel binary workbook): values, formulas (see xlsb_formula; others are saved as their
values and counted), cell formats, column widths, row heights, hidden rows / columns / sheets, merged cells.
Charts, pictures, comments, rules, filters and other extras are not written (use .xlsx for those)."""
import struct
import zipfile

import xlsb_formula
from formatting import DEFAULT_FONT, DEFAULT_SIZE, EPOCH, NUMBER, parse_date

BORDER_CODES = {"thin": 1, "medium": 2, "dashed": 3, "dotted": 4, "thick": 5, "double": 6, "hair": 7,
                "mediumDashed": 8, "dashDot": 9, "mediumDashDot": 10, "dashDotDot": 11, "mediumDashDotDot": 12,
                "slantDashDot": 13}
H_CODES = {"left": 1, "center": 2, "right": 3, "justify": 5}
ERROR_CODES = {"#NULL!": 0x00, "#DIV/0!": 0x07, "#VALUE!": 0x0F, "#REF!": 0x17, "#NAME?": 0x1D, "#NUM!": 0x24,
               "#N/A": 0x2A}
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def rec(rtype, payload=b""):
    out = bytearray()
    if rtype >= 0x80:
        out += bytes([(rtype & 0x7F) | 0x80, rtype >> 7])
    else:
        out.append(rtype)
    n = len(payload)
    while True:
        b = n & 0x7F
        n >>= 7
        out.append(b | (0x80 if n else 0))
        if not n:
            break
    return bytes(out) + payload


def ws(s):
    return struct.pack("<I", len(s)) + s.encode("utf-16-le")


def color(hex_):
    """'#rrggbb' -> BrtColor (RGB); None -> automatic."""
    if not hex_:
        return bytes([0x01, 0, 0, 0, 0, 0, 0, 0xFF])          # automatic (xColorType 0, fValidRGB)
    h = hex_.lstrip("#")
    return bytes([0x05, 0, 0, 0, int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16), 0xFF])


class Styles:
    """Collects fonts / fills / borders / number formats / cell formats while cells are written."""
    def __init__(self):
        from openpyxl.styles.numbers import BUILTIN_FORMATS
        self.builtin = {v: k for k, v in BUILTIN_FORMATS.items()}
        self.fonts = [self.font({})]
        self.fills = [self.fill(0, None), self.fill(17, None)]       # none, gray125: Excel expects both
        self.borders = [self.border({})]
        self.fmts = {}
        self.xfs = [self.xf(0, 0, 0, 0, {})]
        self.index = {None: 0}

    @staticmethod
    def font(st):
        grbit = (0x02 if st.get("italic") else 0) | (0x08 if st.get("strike") else 0)
        size = st.get("size") or DEFAULT_SIZE
        return (struct.pack("<HHHHBBBB", round(size * 20), grbit, 700 if st.get("bold") else 400, 0,
                            1 if st.get("underline") else 0, 2, 1, 0)
                + color(st.get("fg")) + bytes([0]) + ws(st.get("font") or DEFAULT_FONT))

    @staticmethod
    def fill(pattern, bg):
        fore = color(bg) if bg else bytes([0x03, 0x40, 0, 0, 0, 0, 0, 0xFF])     # indexed 64: system colour
        back = bytes([0x03, 0x41, 0, 0, 0, 0, 0, 0xFF])
        return struct.pack("<I", pattern) + fore + back + struct.pack("<I", 0) + b"\0" * 40 + struct.pack("<I", 0)

    @staticmethod
    def border(sides):
        out = bytearray([0])
        for side in ("top", "bottom", "left", "right", "diagonal"):
            style, col = (sides.get(side) or [None, None])[:2]
            code = BORDER_CODES.get(style, 0)
            out += bytes([code, 0]) + (color(col) if code else bytes(8))
        return bytes(out)

    @staticmethod
    def xf(ifmt, ifont, ifill, iborder, st):
        h = H_CODES.get(st.get("align"), 0)
        v = {"top": 0, "center": 1}.get(st.get("valign"), 2)
        rot = st.get("rotate") or 0
        trot = rot if rot >= 0 else 90 - rot           # Excel: 91-180 = 1-90 degrees downwards
        flags = h | (v << 3) | (0x40 if st.get("wrap") else 0) | (0 if st.get("unlocked") else 0x1000)
        flags |= 0x3F << 16                               # xfGrbitAtr: every part is this format's own
        return struct.pack("<HHHHHBBI", 0, ifmt, ifont, ifill, iborder, min(max(trot, 0), 255),
                           min(int(st.get("indent") or 0), 250), flags)

    def _add(self, lst, item):
        if item not in lst:
            lst.append(item)
        return lst.index(item)

    def get(self, st):
        """Style dict -> cell format index."""
        key = repr(sorted(st.items())) if st else None
        if key in self.index:
            return self.index[key]
        code = st.get("numfmt") or "General"
        if code in self.builtin:
            ifmt = self.builtin[code]
        else:
            ifmt = self.fmts.setdefault(code, 164 + len(self.fmts))
        ifont = self._add(self.fonts, self.font(st))
        ifill = self._add(self.fills, self.fill(1, st["bg"])) if st.get("bg") else 0
        iborder = self._add(self.borders, self.border(st.get("border") or {}))
        self.xfs.append(self.xf(ifmt, ifont, ifill, iborder, st))
        self.index[key] = len(self.xfs) - 1
        return self.index[key]

    def part(self):
        out = [rec(0x116)]
        out += [rec(0x267, struct.pack("<I", len(self.fmts)))]
        out += [rec(0x2C, struct.pack("<H", i) + ws(code)) for code, i in self.fmts.items()] + [rec(0x268)]
        out += [rec(0x263, struct.pack("<I", len(self.fonts)))] + [rec(0x2B, f) for f in self.fonts] + [rec(0x264)]
        out += [rec(0x25B, struct.pack("<I", len(self.fills)))] + [rec(0x2D, f) for f in self.fills] + [rec(0x25C)]
        out += [rec(0x265, struct.pack("<I", len(self.borders)))] + [rec(0x2E, b) for b in self.borders]
        out += [rec(0x266)]
        style_xf = struct.pack("<HHHHHBBI", 0xFFFF, 0, 0, 0, 0, 0, 0, (2 << 3) | 0x1000)
        out += [rec(0x272, struct.pack("<I", 1)), rec(0x2F, style_xf), rec(0x273)]
        out += [rec(0x269, struct.pack("<I", len(self.xfs)))] + [rec(0x2F, x) for x in self.xfs] + [rec(0x26A)]
        out += [rec(0x26B, struct.pack("<I", 1)), rec(0x30, struct.pack("<IHBB", 0, 1, 0, 0xFF) + ws("Normal")),
                rec(0x26C)]
        out += [rec(0x1F9, struct.pack("<I", 0)), rec(0x1FA), rec(0x117)]
        return b"".join(out)


class Strings:
    def __init__(self):
        self.items, self.index, self.total = [], {}, 0

    def get(self, s):
        self.total += 1
        if s not in self.index:
            self.index[s] = len(self.items)
            self.items.append(s)
        return self.index[s]

    def part(self):
        return b"".join([rec(0x19F, struct.pack("<II", self.total, len(self.items)))] +
                        [rec(0x13, b"\0" + ws(s)) for s in self.items] + [rec(0x1A0)])


def _number(text):
    """Cell text -> float for numbers / dates, else None."""
    if NUMBER.fullmatch(text):
        return float(text)
    d = parse_date(text)
    if d is not None:
        return (d - EPOCH).total_seconds() / 86400
    return None


def cell_record(r, c, text, ixf, strings, cached=None):
    """One cell -> (record bytes, formula kept?). cached: the formula's calculated value (any type)."""
    head = struct.pack("<II", c, ixf)
    if text.startswith("=") and len(text) > 1:
        rgce = xlsb_formula.encode(text)
        if rgce is not None:
            tail = struct.pack("<H", 0) + struct.pack("<I", len(rgce)) + rgce + struct.pack("<I", 0)
            if isinstance(cached, bool):
                return rec(0x0A, head + bytes([cached]) + tail), True
            if isinstance(cached, (int, float)):
                return rec(0x09, head + struct.pack("<d", float(cached)) + tail), True
            s = "" if cached is None else str(cached)
            if s in ERROR_CODES:
                return rec(0x0B, head + bytes([ERROR_CODES[s]]) + tail), True
            return rec(0x08, head + ws(s) + tail), True
        # can't be written as a formula: its value instead
        text = "" if cached is None else ("TRUE" if cached is True else "FALSE" if cached is False else
                                          xlsb_number_text(cached))
        rec_bytes, _ = cell_record(r, c, text, ixf, strings) if text else (rec(0x01, head), None)
        return rec_bytes, False
    if text == "":
        return rec(0x01, head), None
    if text in ("TRUE", "FALSE"):
        return rec(0x04, head + bytes([text == "TRUE"])), None
    if text in ERROR_CODES:
        return rec(0x03, head + bytes([ERROR_CODES[text]])), None
    x = _number(text)
    if x is not None:
        return rec(0x05, head + struct.pack("<d", x)), None
    return rec(0x07, head + struct.pack("<I", strings.get(text))), None


def xlsb_number_text(v):
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() and abs(v) < 1e15 else f"{v:.15g}"
    return str(v)


def sheet_part(sheet, styles, strings, value_of, selected):
    """value_of(r, c) -> calculated value of a formula cell. Returns (bytes, formulas saved as values)."""
    rows = sheet.rows
    cells = {}
    for r, row in enumerate(rows):
        for c, t in enumerate(row):
            if t != "":
                cells.setdefault(r, {})[c] = t
    for (r, c), st in sheet.styles.items():
        if st:
            cells.setdefault(r, {}).setdefault(c, "")
    max_r = max(list(cells) + [0])
    max_c = max([c for row in cells.values() for c in row] + [0])
    out = [rec(0x81), rec(0x94, struct.pack("<IIII", 0, max_r, 0, max_c))]
    view = struct.pack("<HIIIBBHHHHHI", 0x3DC if selected else 0x39C, 0, 0, 0, 0x40, 0, 0, 100, 0, 0, 0, 0)
    out += [rec(0x85), rec(0x89, view), rec(0x8A), rec(0x86)]
    out += [rec(0x1E5, struct.pack("<iHHHBB", -1, 8, 300, 0, 0, 0))]   # BrtWsFmtInfo: default sizes
    cols = sorted(set(sheet.widths) | set(sheet.hidden_cols))
    if cols:
        out.append(rec(0x186))
        for c in cols:
            w = sheet.widths.get(c, 8.43)
            flags = (0x01 if c in sheet.hidden_cols else 0) | (0x02 if c in sheet.widths else 0)
            out.append(rec(0x3C, struct.pack("<IIIIH", c, c, round(w * 256), 0, flags)))
        out.append(rec(0x187))
    out.append(rec(0x91))
    lost = 0
    for r in sorted(set(cells) | set(sheet.heights) | set(sheet.hidden_rows)):
        row = cells.get(r, {})
        height = sheet.heights.get(r)
        flags = (0x1000 if r in sheet.hidden_rows else 0) | (0x2000 if height else 0)
        span = struct.pack("<III", 1, min(row), max(row)) if row else struct.pack("<I", 0)
        out.append(rec(0x00, struct.pack("<IIHHB", r, 0, round((height or 15) * 20), flags, 0) + span))
        for c in sorted(row):
            t = row[c]
            ixf = styles.get(sheet.styles.get((r, c)))
            cached = value_of(r, c) if t.startswith("=") else None
            data, kept = cell_record(r, c, t, ixf, strings, cached)
            if kept is False:
                lost += 1
            out.append(data)
    out.append(rec(0x92))
    if sheet.merges:
        out.append(rec(0xB1, struct.pack("<I", len(sheet.merges))))
        out += [rec(0xB0, struct.pack("<IIII", r1, r2, c1, c2)) for r1, c1, r2, c2 in sheet.merges]
        out.append(rec(0xB2))
    out.append(rec(0x82))
    return b"".join(out), lost


def write(dst, app, sheets):
    """Write every sheet to dst (.xlsb). Returns how many formulas were saved as values."""
    from formula import XLError
    eng = app.engine
    styles, strings = Styles(), Strings()
    parts, lost = [], 0
    names = list(sheets)
    first_shown = next((i for i, n in enumerate(names) if not sheets[n].hidden), 0)
    for i, name in enumerate(names):
        def value_of(r, c, name=name):
            v = eng.cell(name, r, c)
            if isinstance(v, XLError):
                return str(v)
            return v
        data, n = sheet_part(sheets[name], styles, strings, value_of, i == first_shown)
        parts.append(data)
        lost += n
    book = [rec(0x83), rec(0x99, struct.pack("<II", 0, 0) + ws(""))]
    book += [rec(0x87), rec(0x9E, struct.pack("<iiIIIIIB", 0, 0, 20000, 12000, 600, 0, first_shown, 0x78)), rec(0x88)]
    book.append(rec(0x8F))
    for i, name in enumerate(names):
        state = 1 if sheets[name].hidden else 0
        book.append(rec(0x9C, struct.pack("<II", state, i + 1) + ws(f"rId{i + 1}") + ws(name)))
    book += [rec(0x90), rec(0x84)]
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="bin" ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.main"/>'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          + "".join(f'<Override PartName="/xl/worksheets/sheet{i + 1}.bin" '
                    'ContentType="application/vnd.ms-excel.worksheet"/>' for i in range(len(names)))
          + '<Override PartName="/xl/styles.bin" ContentType="application/vnd.ms-excel.styles"/>'
          '<Override PartName="/xl/sharedStrings.bin" ContentType="application/vnd.ms-excel.sharedStrings"/>'
          '</Types>')
    root_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 f'<Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.bin"/></Relationships>')
    wb_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
               + "".join(f'<Relationship Id="rId{i + 1}" Type="{REL}/worksheet" Target="worksheets/sheet{i + 1}.bin"/>'
                         for i in range(len(names)))
               + f'<Relationship Id="rId{len(names) + 1}" Type="{REL}/styles" Target="styles.bin"/>'
               f'<Relationship Id="rId{len(names) + 2}" Type="{REL}/sharedStrings" Target="sharedStrings.bin"/>'
               '</Relationships>')
    import os
    import tempfile
    fd, tmp = tempfile.mkstemp(suffix=".xlsb", dir=os.path.dirname(os.path.abspath(dst)))
    os.close(fd)
    try:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct)
            z.writestr("_rels/.rels", root_rels)
            z.writestr("xl/workbook.bin", b"".join(book))
            z.writestr("xl/_rels/workbook.bin.rels", wb_rels)
            for i, data in enumerate(parts):
                z.writestr(f"xl/worksheets/sheet{i + 1}.bin", data)
            z.writestr("xl/styles.bin", styles.part())
            z.writestr("xl/sharedStrings.bin", strings.part())
        os.replace(tmp, dst)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)
    return lost
