"""Formulas from old .xls files (xlrd only gives their results). Reads the BIFF8 workbook stream and turns
each FORMULA record's tokens back into Excel formula text. Every decoded formula is checked against the
result Excel saved; when they differ (or a token isn't understood) the cell keeps the saved value."""
import struct

from formula import col_letters, quote_sheet

# BIFF8 function numbers -> (name, fixed argument count or None)
FUNCS = {
    0: ("COUNT", None), 1: ("IF", None), 2: ("ISNA", 1), 3: ("ISERROR", 1), 4: ("SUM", None), 5: ("AVERAGE", None),
    6: ("MIN", None), 7: ("MAX", None), 8: ("ROW", None), 9: ("COLUMN", None), 10: ("NA", 0), 11: ("NPV", None),
    12: ("STDEV", None), 13: ("DOLLAR", None), 14: ("FIXED", None), 15: ("SIN", 1), 16: ("COS", 1), 17: ("TAN", 1),
    18: ("ATAN", 1), 19: ("PI", 0), 20: ("SQRT", 1), 21: ("EXP", 1), 22: ("LN", 1), 23: ("LOG10", 1),
    24: ("ABS", 1), 25: ("INT", 1), 26: ("SIGN", 1), 27: ("ROUND", 2), 28: ("LOOKUP", None), 29: ("INDEX", None),
    30: ("REPT", 2), 31: ("MID", 3), 32: ("LEN", 1), 33: ("VALUE", 1), 34: ("TRUE", 0), 35: ("FALSE", 0),
    36: ("AND", None), 37: ("OR", None), 38: ("NOT", 1), 39: ("MOD", 2), 46: ("VAR", None), 48: ("TEXT", 2),
    56: ("PV", None), 57: ("FV", None), 58: ("NPER", None), 59: ("PMT", None), 60: ("RATE", None), 62: ("IRR", None),
    63: ("RAND", 0), 64: ("MATCH", None), 65: ("DATE", 3), 66: ("TIME", 3), 67: ("DAY", 1), 68: ("MONTH", 1),
    69: ("YEAR", 1), 70: ("WEEKDAY", None), 71: ("HOUR", 1), 72: ("MINUTE", 1), 73: ("SECOND", 1), 74: ("NOW", 0),
    76: ("ROWS", 1), 77: ("COLUMNS", 1), 78: ("OFFSET", None), 82: ("SEARCH", None), 83: ("TRANSPOSE", 1),
    86: ("TYPE", 1), 97: ("ATAN2", 2), 98: ("ASIN", 1), 99: ("ACOS", 1), 100: ("CHOOSE", None),
    101: ("HLOOKUP", None), 102: ("VLOOKUP", None), 105: ("ISREF", 1), 109: ("LOG", None), 111: ("CHAR", 1),
    112: ("LOWER", 1), 113: ("UPPER", 1), 114: ("PROPER", 1), 115: ("LEFT", None), 116: ("RIGHT", None),
    117: ("EXACT", 2), 118: ("TRIM", 1), 119: ("REPLACE", 4), 120: ("SUBSTITUTE", None), 121: ("CODE", 1),
    124: ("FIND", None), 125: ("CELL", None), 126: ("ISERR", 1), 127: ("ISTEXT", 1), 128: ("ISNUMBER", 1),
    129: ("ISBLANK", 1), 130: ("T", 1), 131: ("N", 1), 140: ("DATEVALUE", 1), 141: ("TIMEVALUE", 1),
    142: ("SLN", 3), 148: ("INDIRECT", None), 162: ("CLEAN", 1), 163: ("MDETERM", 1), 165: ("MMULT", 2),
    167: ("IPMT", None), 168: ("PPMT", None), 169: ("COUNTA", None), 183: ("PRODUCT", None), 184: ("FACT", 1),
    190: ("ISNONTEXT", 1), 193: ("STDEVP", None), 194: ("VARP", None), 197: ("TRUNC", None), 198: ("ISLOGICAL", 1),
    212: ("ROUNDUP", 2), 213: ("ROUNDDOWN", 2), 216: ("RANK", None), 219: ("ADDRESS", None), 220: ("DAYS360", None),
    221: ("TODAY", 0), 227: ("MEDIAN", None), 228: ("SUMPRODUCT", None), 252: ("FREQUENCY", 2),
    269: ("AVEDEV", None), 276: ("COMBIN", 2), 279: ("EVEN", 1), 285: ("FLOOR", 2), 288: ("CEILING", 2),
    298: ("ODD", 1), 325: ("LARGE", 2), 326: ("SMALL", 2), 327: ("QUARTILE", 2), 328: ("PERCENTILE", 2),
    330: ("MODE", None), 336: ("CONCATENATE", None), 337: ("POWER", 2), 342: ("RADIANS", 1), 343: ("DEGREES", 1),
    344: ("SUBTOTAL", None), 345: ("SUMIF", None), 346: ("COUNTIF", 2), 347: ("COUNTBLANK", 1),
    359: ("HYPERLINK", None), 362: ("MAXA", None), 363: ("MINA", None),
}
OPS = {0x03: "+", 0x04: "-", 0x05: "*", 0x06: "/", 0x07: "^", 0x08: "&", 0x09: "<", 0x0A: "<=", 0x0B: "=",
       0x0C: ">=", 0x0D: ">", 0x0E: "<>", 0x10: ",", 0x11: ":"}
ERRS = {0x00: "#NULL!", 0x07: "#DIV/0!", 0x0F: "#VALUE!", 0x17: "#REF!", 0x1D: "#NAME?", 0x24: "#NUM!", 0x2A: "#N/A"}


class Unsupported(Exception):
    pass


def _cell(row, colfield, base=None):
    """Reference text from BIFF8 row / column fields (bit 15 = row relative, bit 14 = column relative).
    base=(r, c): RefN / AreaN in shared formulas, where relative parts are offsets."""
    col, rrel, crel = colfield & 0x3FFF, colfield & 0x8000, colfield & 0x4000
    if base is not None:
        if rrel:
            row = base[0] + (row - 0x10000 if row & 0x8000 else row)
        if crel:
            col = base[1] + ((col & 0xFF) - 0x100 if col & 0x80 else col & 0xFF)
    if row < 0 or col < 0:
        raise Unsupported("ref")
    return f"{'' if crel else '$'}{col_letters(col)}{'' if rrel else '$'}{row + 1}"


def decode(rgce, book, base=None, names=()):
    """Token bytes -> formula text without '='. base: position for relative (shared-formula) tokens."""
    st, i = [], 0
    while i < len(rgce):
        t = rgce[i]
        i += 1
        if t in OPS:
            b, a = st.pop(), st.pop()
            st.append(f"{a}{OPS[t]}{b}")
        elif t == 0x0F:
            b, a = st.pop(), st.pop()
            st.append(f"{a} {b}")
        elif t == 0x12:
            st.append("+" + st.pop())
        elif t == 0x13:
            st.append("-" + st.pop())
        elif t == 0x14:
            st.append(st.pop() + "%")
        elif t == 0x15:
            st.append(f"({st.pop()})")
        elif t == 0x16:
            st.append("")
        elif t == 0x17:
            n, flags = rgce[i], rgce[i + 1]
            i += 2
            if flags & 1:
                s = rgce[i:i + 2 * n].decode("utf-16-le")
                i += 2 * n
            else:
                s = rgce[i:i + n].decode("latin-1")
                i += n
            st.append('"' + s.replace('"', '""') + '"')
        elif t == 0x19:                                      # attributes: spaces, IF / CHOOSE jumps, SUM
            kind = rgce[i]
            if kind & 0x04:                                  # tAttrChoose: jump table
                n = struct.unpack_from("<H", rgce, i + 1)[0]
                i += 3 + 2 * (n + 1)
                continue
            if kind & 0x10:                                  # tAttrSum = SUM(top)
                st.append(f"SUM({st.pop()})")
            i += 3
        elif t == 0x1C:
            st.append(ERRS.get(rgce[i], "#N/A"))
            i += 1
        elif t == 0x1D:
            st.append("TRUE" if rgce[i] else "FALSE")
            i += 1
        elif t == 0x1E:
            st.append(str(struct.unpack_from("<H", rgce, i)[0]))
            i += 2
        elif t == 0x1F:
            v = struct.unpack_from("<d", rgce, i)[0]
            st.append(str(int(v)) if v.is_integer() and abs(v) < 1e15 else repr(v))
            i += 8
        elif t in (0x21, 0x41, 0x61):
            fid = struct.unpack_from("<H", rgce, i)[0]
            i += 2
            name, argc = FUNCS.get(fid, (None, None))
            if name is None or argc is None:
                raise Unsupported(f"function {fid}")
            args = [st.pop() for _ in range(argc)][::-1]
            st.append(f"{name}({','.join(args)})")
        elif t in (0x22, 0x42, 0x62):
            argc, fid = rgce[i] & 0x7F, struct.unpack_from("<H", rgce, i + 1)[0] & 0x7FFF
            i += 3
            name = FUNCS.get(fid, (None,))[0]
            if name is None:
                raise Unsupported(f"function {fid}")
            args = [st.pop() for _ in range(argc)][::-1]
            st.append(f"{name}({','.join(args)})")
        elif t in (0x23, 0x43, 0x63):
            k = struct.unpack_from("<H", rgce, i)[0]
            i += 4
            if not 0 < k <= len(names) or not names[k - 1]:
                raise Unsupported("name")
            st.append(names[k - 1])
        elif t in (0x24, 0x44, 0x64):
            row, col = struct.unpack_from("<HH", rgce, i)
            i += 4
            st.append(_cell(row, col))
        elif t in (0x25, 0x45, 0x65):
            r1, r2, c1, c2 = struct.unpack_from("<HHHH", rgce, i)
            i += 8
            st.append(f"{_cell(r1, c1)}:{_cell(r2, c2)}")
        elif t in (0x2C, 0x4C, 0x6C):
            row, col = struct.unpack_from("<HH", rgce, i)
            i += 4
            st.append(_cell(row, col, base))
        elif t in (0x2D, 0x4D, 0x6D):
            r1, r2, c1, c2 = struct.unpack_from("<HHHH", rgce, i)
            i += 8
            st.append(f"{_cell(r1, c1, base)}:{_cell(r2, c2, base)}")
        elif t in (0x2A, 0x4A, 0x6A):
            i += 4
            st.append("#REF!")
        elif t in (0x2B, 0x4B, 0x6B):
            i += 8
            st.append("#REF!")
        elif t in (0x3A, 0x5A, 0x7A):
            ixti, row, col = struct.unpack_from("<HHH", rgce, i)
            i += 6
            st.append(f"{book.sheet_ref(ixti)}!{_cell(row, col)}")
        elif t in (0x3B, 0x5B, 0x7B):
            ixti, r1, r2, c1, c2 = struct.unpack_from("<HHHHH", rgce, i)
            i += 10
            st.append(f"{book.sheet_ref(ixti)}!{_cell(r1, c1)}:{_cell(r2, c2)}")
        elif t in (0x26, 0x46, 0x66, 0x27, 0x47, 0x67):        # mem area / mem error: the area follows
            i += 6
        elif t in (0x29, 0x49, 0x69):                        # mem func: the expression follows
            i += 2
        else:
            raise Unsupported(f"token {t:#x}")
    if len(st) != 1:
        raise Unsupported("stack")
    return st[0]


class Book:
    def __init__(self, data):
        self.data = data
        self.sheets, self.externs, self.supbooks, self.names = [], [], [], []

    def records(self, pos=0):
        d = self.data
        while pos + 4 <= len(d):
            rtype, size = struct.unpack_from("<HH", d, pos)
            yield pos, rtype, d[pos + 4:pos + 4 + size]
            pos += 4 + size

    def sheet_ref(self, ixti):
        if ixti >= len(self.externs):
            raise Unsupported("extern")
        sup, first, last = self.externs[ixti]
        if sup >= len(self.supbooks) or not self.supbooks[sup] or first >= len(self.sheets) or first != last:
            raise Unsupported("external workbook")
        return quote_sheet(self.sheets[first][0])

    def read_globals(self):
        for _pos, rtype, p in self.records():
            if rtype == 0x0085:                              # BOUNDSHEET
                offset, _vis, kind = struct.unpack_from("<IBB", p, 0)
                n, flags = p[6], p[7]
                name = p[8:8 + 2 * n].decode("utf-16-le") if flags & 1 else p[8:8 + n].decode("latin-1")
                if kind == 0:                                # worksheets only
                    self.sheets.append((name, offset))
            elif rtype == 0x01AE:                            # SUPBOOK: internal one is 0x0401
                self.supbooks.append(len(p) >= 4 and struct.unpack_from("<H", p, 2)[0] == 0x0401)
            elif rtype == 0x0017:                            # EXTERNSHEET
                (n,) = struct.unpack_from("<H", p, 0)
                self.externs = [struct.unpack_from("<HHH", p, 2 + 6 * k) for k in range(n)]
            elif rtype == 0x0018:                            # NAME (defined names), for ptgName
                try:
                    n = p[3]
                    flags = p[14]
                    raw = p[15:15 + (2 * n if flags & 1 else n)]
                    self.names.append(raw.decode("utf-16-le" if flags & 1 else "latin-1") if not p[0] & 0x20 else "")
                except IndexError:
                    self.names.append("")
            elif rtype == 0x000A:                            # EOF of the globals
                break

    def formulas(self, sheet_index):
        """{(r, c): formula text} for one worksheet."""
        _, offset = self.sheets[sheet_index]
        out, shared = {}, {}
        pending = []
        for _pos, rtype, p in self.records(offset):
            if rtype == 0x0006 and len(p) >= 22:             # FORMULA
                r, c = struct.unpack_from("<HH", p, 0)
                cce = struct.unpack_from("<H", p, 20)[0]
                rgce = p[22:22 + cce]
                if rgce[:1] == b"\x01":                      # ptgExp: part of a shared formula
                    pending.append((r, c, struct.unpack_from("<HH", rgce, 1)))
                    continue
                try:
                    out[(r, c)] = "=" + decode(rgce, self, (r, c), self.names)
                except (Unsupported, IndexError, struct.error, UnicodeDecodeError):
                    pass
            elif rtype == 0x04BC:                            # SHRFMLA: rows/cols, then the tokens
                r1, r2, c1, c2 = struct.unpack_from("<HHBB", p, 0)
                cce = struct.unpack_from("<H", p, 8)[0]
                shared[(r1, c1)] = (r1, r2, c1, c2, p[10:10 + cce])
            elif rtype == 0x000A:
                break
        for r, c, anchor in pending:
            sh = shared.get(anchor)
            if sh is None:
                continue
            try:
                out[(r, c)] = "=" + decode(sh[4], self, (r, c), self.names)
            except (Unsupported, IndexError, struct.error, UnicodeDecodeError):
                pass
        return out


def read(path):
    """[{(r, c): '=formula'}] per worksheet, in order (empty when the file can't be read this way)."""
    from simplecore.office_crypto import CompoundFile, NotEncrypted
    try:
        with open(path, "rb") as f:
            cf = CompoundFile(f.read())
        try:
            data = cf.stream("Workbook")
        except KeyError:
            data = cf.stream("Book")
    except (OSError, KeyError, NotEncrypted, struct.error):
        return []
    book = Book(data)
    try:
        book.read_globals()
        return [book.formulas(i) for i in range(len(book.sheets))]
    except (struct.error, IndexError):
        return []


def apply(sheets, path):
    """Put decoded formulas into loaded .xls sheets where they reproduce Excel's saved result."""
    from formula import Engine
    found = read(path)
    if not found:
        return 0
    eng = Engine()
    eng.set_book(sheets)
    names = list(sheets)
    used = 0
    for i, formulas in enumerate(found):
        if i >= len(names):
            break
        s = sheets[names[i]]
        for (r, c), f in formulas.items():
            if r >= len(s.rows) or c >= len(s.rows[r]):
                continue
            saved = s.rows[r][c]
            s.rows[r][c] = f
            eng.clear()
            got = eng.display(s, r, c)
            if saved != "" and not same(got, saved):   # (no saved result, e.g. not written by Excel: keep it)
                s.rows[r][c] = saved                         # couldn't reproduce it: keep the value
                continue
            s.cache[(r, c)] = saved
            used += 1
    return used


def same(got, saved):
    if got == saved:
        return True
    try:
        a, b = float(got), float(saved)
    except ValueError:
        return got.upper() in ("TRUE", "FALSE") and got.upper() == saved.upper()
    return abs(a - b) <= 1e-9 * max(1.0, abs(a), abs(b))
