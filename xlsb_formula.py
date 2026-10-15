"""BIFF12 (.xlsb) formulas: parsed tokens (ptgs, reverse Polish) <-> formula text.

decode() reads what Excel wrote; anything it doesn't know returns None (the caller keeps the saved result).
encode() writes cell formulas: numbers, text, TRUE/FALSE, errors, cell and range references on the same sheet,
operators, parentheses and the classic function set (xlrd's table); None for other formulas (cross-sheet,
names, newer functions...), which are then saved as values.
"""
import re
import struct

from xlrd.formula import func_defs

from formula import col_index, col_letters, lex

MAXROW, MAXCOL = 1048575, 16383
BINOPS = {0x03: "+", 0x04: "-", 0x05: "*", 0x06: "/", 0x07: "^", 0x08: "&", 0x09: "<", 0x0A: "<=", 0x0B: "=",
          0x0C: ">=", 0x0D: ">", 0x0E: "<>", 0x0F: " ", 0x10: ",", 0x11: ":"}
OPCODES = {v: k for k, v in BINOPS.items() if k <= 0x0E}
ERRORS = {0x00: "#NULL!", 0x07: "#DIV/0!", 0x0F: "#VALUE!", 0x17: "#REF!", 0x1D: "#NAME?", 0x24: "#NUM!", 0x2A: "#N/A",
          0x2B: "#GETTING_DATA"}
ERROR_CODES = {v: k for k, v in ERRORS.items()}
FUNCS = {v[0]: k for k, v in func_defs.items()}


class Unsupported(Exception):
    pass


# ---------- decoding ----------

def _col_text(col_field):
    col = col_field & 0x3FFF
    return col, not col_field & 0x4000, not col_field & 0x8000      # col, col absolute, row absolute


def _cell(r, c, rabs, cabs):
    return f"{'$' if cabs else ''}{col_letters(c)}{'$' if rabs else ''}{r + 1}"


def _area(r1, c1, r2, c2, f1, f2):
    """f = (row abs, col abs) for each corner."""
    if r1 == 0 and r2 == MAXROW:          # whole columns: A:C
        return f"{'$' if f1[1] else ''}{col_letters(c1)}:{'$' if f2[1] else ''}{col_letters(c2)}"
    if c1 == 0 and c2 == MAXCOL:          # whole rows: 2:5
        return f"{'$' if f1[0] else ''}{r1 + 1}:{'$' if f2[0] else ''}{r2 + 1}"
    return _cell(r1, c1, *f1) + ":" + _cell(r2, c2, *f2)


def _sheet_prefix(sheets, ixti):
    if not sheets or ixti >= len(sheets) or sheets[ixti] is None:
        raise Unsupported("sheet")
    name = sheets[ixti]
    from formula import quote_sheet
    return quote_sheet(name) + "!"


def _signed(n, bits):
    return n - (1 << bits) if n & (1 << (bits - 1)) else n


def decode(rgce, base=(0, 0), sheets=None):
    """Formula tokens -> text without the leading '=' (None if something isn't supported).
    base: the cell (for shared-formula relative refs); sheets: XTI index -> sheet name for 3-D refs."""
    try:
        return _decode(rgce, base, sheets)
    except (Unsupported, IndexError, struct.error, KeyError):
        return None


def _decode(rgce, base, sheets):
    stack, i, n = [], 0, len(rgce)
    while i < n:
        ptg = rgce[i]
        i += 1
        k = ptg if ptg < 0x20 else ((ptg - 0x20) & 0x1F) | 0x20
        if k in BINOPS and ptg < 0x20:
            b, a = stack.pop(), stack.pop()
            stack.append(f"{a}{BINOPS[k]}{b}")
        elif k == 0x12:
            stack.append("+" + stack.pop())
        elif k == 0x13:
            stack.append("-" + stack.pop())
        elif k == 0x14:
            stack.append(stack.pop() + "%")
        elif k == 0x15:
            stack.append(f"({stack.pop()})")
        elif k == 0x16:
            stack.append("")
        elif k == 0x17:
            (cch,) = struct.unpack_from("<H", rgce, i)
            s = rgce[i + 2:i + 2 + cch * 2].decode("utf-16-le")
            i += 2 + cch * 2
            stack.append('"' + s.replace('"', '""') + '"')
        elif k == 0x19:                                     # PtgAttr
            grbit = rgce[i]
            (data,) = struct.unpack_from("<H", rgce, i + 1)
            i += 3
            if grbit & 0x04:                                # tAttrChoose: jump table
                i += (data + 1) * 2
            if grbit & 0x10:                                # tAttrSum: SUM of one argument
                stack.append(f"SUM({stack.pop()})")
        elif k == 0x1C:
            stack.append(ERRORS.get(rgce[i], "#N/A"))
            i += 1
        elif k == 0x1D:
            stack.append("TRUE" if rgce[i] else "FALSE")
            i += 1
        elif k == 0x1E:
            stack.append(str(struct.unpack_from("<H", rgce, i)[0]))
            i += 2
        elif k == 0x1F:
            (x,) = struct.unpack_from("<d", rgce, i)
            i += 8
            stack.append(str(int(x)) if x.is_integer() and abs(x) < 1e15 else repr(x))
        elif k == 0x21 or k == 0x22:                        # PtgFunc / PtgFuncVar
            if k == 0x21:
                (tab,) = struct.unpack_from("<H", rgce, i)
                i += 2
                argc = func_defs[tab][1]
            else:
                argc = rgce[i]
                (tab,) = struct.unpack_from("<H", rgce, i + 1)
                i += 3
                if tab & 0x8000:
                    raise Unsupported("command")
            if tab not in func_defs:
                raise Unsupported("function")
            args = stack[len(stack) - argc:] if argc else []
            del stack[len(stack) - argc:]
            stack.append(f"{func_defs[tab][0]}({','.join(args)})")
        elif k in (0x26, 0x27, 0x28, 0x29):                 # PtgMemArea / MemErr / MemNoMem / MemFunc: markers
            i += 6 if k in (0x26, 0x27) else 2
        elif k == 0x24 or k == 0x2C:                        # PtgRef / PtgRefN
            (row,) = struct.unpack_from("<I", rgce, i)
            (cf,) = struct.unpack_from("<H", rgce, i + 4)
            i += 6
            c, cabs, rabs = _col_text(cf)
            if k == 0x2C:
                row = row if rabs else base[0] + _signed(row, 32)
                c = c if cabs else base[1] + _signed(c, 14)
            stack.append(_cell(row, c, rabs, cabs))
        elif k == 0x25 or k == 0x2D:                        # PtgArea / PtgAreaN
            r1, r2, cf1, cf2 = struct.unpack_from("<IIHH", rgce, i)
            i += 12
            c1, ca1, ra1 = _col_text(cf1)
            c2, ca2, ra2 = _col_text(cf2)
            if k == 0x2D:
                r1 = r1 if ra1 else base[0] + _signed(r1, 32)
                r2 = r2 if ra2 else base[0] + _signed(r2, 32)
                c1 = c1 if ca1 else base[1] + _signed(c1, 14)
                c2 = c2 if ca2 else base[1] + _signed(c2, 14)
            stack.append(_area(r1, c1, r2, c2, (ra1, ca1), (ra2, ca2)))
        elif k in (0x2A, 0x2B):                             # PtgRefErr / PtgAreaErr
            i += 6 if k == 0x2A else 12
            stack.append("#REF!")
        elif k == 0x3A:                                     # PtgRef3d
            (ixti, row, cf) = struct.unpack_from("<HIH", rgce, i)
            i += 8
            c, cabs, rabs = _col_text(cf)
            stack.append(_sheet_prefix(sheets, ixti) + _cell(row, c, rabs, cabs))
        elif k == 0x3B:                                     # PtgArea3d
            ixti, r1, r2, cf1, cf2 = struct.unpack_from("<HIIHH", rgce, i)
            i += 14
            c1, ca1, ra1 = _col_text(cf1)
            c2, ca2, ra2 = _col_text(cf2)
            stack.append(_sheet_prefix(sheets, ixti) + _area(r1, c1, r2, c2, (ra1, ca1), (ra2, ca2)))
        elif k in (0x3C, 0x3D):
            i += 8 if k == 0x3C else 14
            stack.append("#REF!")
        else:                                               # names, arrays, tables, newer functions...
            raise Unsupported(hex(ptg))
    if len(stack) != 1:
        raise Unsupported("stack")
    return stack[0]


def shared_ref(rgce):
    """A formula that is only PtgExp (part of a shared formula) -> the row of its first cell, else None."""
    if len(rgce) == 5 and rgce[0] == 0x01:
        return struct.unpack_from("<I", rgce, 1)[0]
    return None


# ---------- encoding ----------

class _Enc:
    def __init__(self, text):
        self.toks = [(k, t) for k, t in lex(text) if k != "ws"]
        self.i = 0
        self.out = bytearray()

    def peek(self):
        return self.toks[self.i] if self.i < len(self.toks) else (None, None)

    def take(self):
        t = self.peek()
        self.i += 1
        return t

    def run(self):
        self.compare("V")
        if self.i != len(self.toks):
            raise Unsupported("trailing")
        return bytes(self.out)

    def binary(self, ops, sub, cls):
        def more():
            return self.peek()[0] == "op" and self.peek()[1] in ops
        start_i, start_n = self.i, len(self.out)
        sub(cls)
        if cls == "R" and more():             # operands of an operator are values: redo the first one
            self.i = start_i
            del self.out[start_n:]
            cls = "V"
            sub(cls)
        while more():
            op = self.take()[1]
            sub(cls)
            self.out.append(OPCODES[op])

    def compare(self, cls):
        self.binary(("=", "<>", "<", ">", "<=", ">="), self.concat, cls)

    def concat(self, cls):
        self.binary(("&",), self.additive, cls)

    def additive(self, cls):
        self.binary(("+", "-"), self.term, cls)

    def term(self, cls):
        self.binary(("*", "/"), self.power, cls)

    def power(self, cls):
        self.binary(("^",), self.unary, cls)

    def unary(self, cls):
        if self.peek() in (("op", "-"), ("op", "+")):
            op = self.take()[1]
            self.unary("V" if cls == "R" else cls)
            self.out.append(0x13 if op == "-" else 0x12)
            return
        start_i, start_n = self.i, len(self.out)
        self.primary(cls)
        if cls == "R" and self.peek() == ("op", "%"):
            self.i = start_i
            del self.out[start_n:]
            self.primary("V")
        while self.peek() == ("op", "%"):
            self.take()
            self.out.append(0x14)

    def primary(self, cls):
        kind, tok = self.take()
        if kind == "num":
            x = float(tok)
            if x.is_integer() and 0 <= x <= 65535:
                self.out += b"\x1E" + struct.pack("<H", int(x))
            else:
                self.out += b"\x1F" + struct.pack("<d", x)
        elif kind == "str":
            s = tok[1:-1].replace('""', '"')
            if len(s) > 255:
                raise Unsupported("long text")
            self.out += b"\x17" + struct.pack("<H", len(s)) + s.encode("utf-16-le")
        elif kind == "bool":
            self.out += b"\x1D" + bytes([tok.upper() == "TRUE"])
        elif kind == "err":
            if tok.upper() not in ERROR_CODES:
                raise Unsupported("error")
            self.out += b"\x1C" + bytes([ERROR_CODES[tok.upper()]])
        elif kind in ("ref", "cols", "rows"):
            self.reference(kind, tok, cls)
        elif kind == "func":
            self.function(tok, cls)
        elif (kind, tok) == ("op", "("):
            self.compare(cls)
            if self.take() != ("op", ")"):
                raise Unsupported("paren")
            self.out.append(0x15)
        else:
            raise Unsupported(kind)

    def reference(self, kind, tok, cls):
        if "!" in tok:
            raise Unsupported("sheet")
        add = {"R": 0x00, "V": 0x20, "A": 0x40}[cls]
        a, _, b = tok.partition(":")
        if kind == "ref" and not b:
            r, c, rabs, cabs = _parse_cell(a)
            self.out += bytes([0x24 + add]) + struct.pack("<IH", r, _col_field(c, rabs, cabs))
            return
        if kind == "ref":
            (r1, c1, ra1, ca1), (r2, c2, ra2, ca2) = _parse_cell(a), _parse_cell(b)
        elif kind == "cols":
            r1, r2, ra1, ra2 = 0, MAXROW, True, True
            c1, c2, ca1, ca2 = col_index(a.strip("$")), col_index(b.strip("$")), a.startswith("$"), b.startswith("$")
        else:
            c1, c2, ca1, ca2 = 0, MAXCOL, True, True
            r1, r2, ra1, ra2 = int(a.strip("$")) - 1, int(b.strip("$")) - 1, a.startswith("$"), b.startswith("$")
        self.out += bytes([0x25 + add]) + struct.pack("<IIHH", r1, r2, _col_field(c1, ra1, ca1), _col_field(c2, ra2, ca2))

    def function(self, tok, cls):
        name = tok.upper()
        if name not in FUNCS:
            raise Unsupported(name)
        tab = FUNCS[name]
        _, lo, hi, *_rest = func_defs[tab]
        ret, arg_cls = func_defs[tab][5], func_defs[tab][6]
        self.take()                                          # (
        argc = 0
        if self.peek() != ("op", ")"):
            while True:
                if self.peek()[0] == "op" and self.peek()[1] in (",", ";", ")"):
                    self.out.append(0x16)
                else:
                    want = arg_cls[argc] if argc < len(arg_cls) else (arg_cls[-1] if arg_cls else "V")
                    self.compare(want if want in "RVA" else "V")
                argc += 1
                if self.peek()[0] == "op" and self.peek()[1] in (",", ";"):
                    self.take()
                    continue
                break
        if self.take() != ("op", ")") or not lo <= argc <= hi:
            raise Unsupported("arguments")
        add = 0x00 if (ret == "R" and cls == "R") else 0x40 if cls == "A" else 0x20
        if lo == hi:
            self.out += bytes([0x21 + add]) + struct.pack("<H", tab)
        else:
            self.out += bytes([0x22 + add, argc]) + struct.pack("<H", tab)


def _parse_cell(text):
    m = re.fullmatch(r"(\$?)([A-Za-z]{1,3})(\$?)(\d+)", text)
    if not m:
        raise Unsupported("cell")
    return int(m.group(4)) - 1, col_index(m.group(2)), bool(m.group(3)), bool(m.group(1))


def _col_field(c, rabs, cabs):
    return (c & 0x3FFF) | (0 if cabs else 0x4000) | (0 if rabs else 0x8000)


def encode(text):
    """'=SUM(A1:B2)*2' -> rgce bytes, or None when it can't be written (saved as its value instead)."""
    body = text[1:] if text.startswith("=") else text
    try:
        return _Enc(body).run()
    except (Unsupported, KeyError, ValueError, struct.error, AttributeError):
        return None
