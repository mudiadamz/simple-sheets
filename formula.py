"""Small Excel formula engine: parse and evaluate common formulas over the loaded book."""
import datetime as dt
import math
import random
import re
import statistics
from decimal import ROUND_DOWN, ROUND_HALF_UP, ROUND_UP, Decimal

from formatting import EPOCH, fmt_number, parse_date, parse_date_text, parse_number, render

MAXR, MAXC = 1048576, 16384


class XLError(str):
    """An Excel error value such as #DIV/0!."""


class Serial(float):
    """A number that is a date (Excel serial day); shown as a date."""


class FormulaError(Exception):
    def __init__(self, err):
        super().__init__(err)
        self.err = err


DIV0, NA, NAME, NUM, REF, VALUE, CIRC = (XLError(e) for e in (
    "#DIV/0!", "#N/A", "#NAME?", "#NUM!", "#REF!", "#VALUE!", "#CIRC!"))
SPILL, CALC = XLError("#SPILL!"), XLError("#CALC!")
ERRORS = {e: e for e in (DIV0, NA, NAME, NUM, REF, VALUE, CIRC, SPILL, CALC, XLError("#NULL!"))}
MISSING = object()   # an empty argument, e.g. IF(A1,,1)


# ---------- references ----------

def col_index(letters):
    n = 0
    for ch in letters.upper():
        n = n * 26 + ord(ch) - 64
    return n - 1


def col_letters(i):
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


SHEET = r"(?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!"
CELL = r"\$?[A-Za-z]{1,3}\$?\d+"
TOKEN = re.compile(rf"""
  (?P<ws>\s+)
| (?P<str>"(?:[^"]|"")*")
| (?P<err>\#(?:DIV/0!|N/A|NAME\?|NULL!|NUM!|REF!|VALUE!|CIRC!|SPILL!|CALC!))
| (?P<ext>'\[[^\]]+\][^']*'!{CELL}(?::{CELL})?|\[[^\]]+\][A-Za-z_][\w.]*!{CELL}(?::{CELL})?)
| (?P<table>(?:[A-Za-z_][\w.]*)?\[(?:[^\[\]]|\[[^\[\]]*\])*\])
| (?P<array>\{{[^{{}}]*\}})
| (?P<spillref>(?:{SHEET})?{CELL}\#)
| (?P<func>(?:_xlfn\.|_xlws\.)?[A-Za-z_][\w.]*)(?=\s*\()
| (?P<ref>(?:{SHEET})?{CELL}(?::{CELL})?)(?![\w(])
| (?P<cols>(?:{SHEET})?\$?[A-Za-z]{{1,3}}:\$?[A-Za-z]{{1,3}})(?![\w(])
| (?P<rows>(?:{SHEET})?\$?\d+:\$?\d+)
| (?P<num>(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?)
| (?P<bool>TRUE|FALSE)(?![\w(])
| (?P<name>[A-Za-z_][\w.]*)
| (?P<op><>|<=|>=|[-+*/^&=<>%(),;])
""", re.X | re.I)


def lex(s):
    """-> [(kind, text)], unknown characters become ('bad', ch)."""
    out, pos = [], 0
    while pos < len(s):
        m = TOKEN.match(s, pos)
        if not m:
            out.append(("bad", s[pos]))
            pos += 1
            continue
        out.append((m.lastgroup, m.group()))
        pos = m.end()
    return out


def split_sheet(tok):
    if "!" not in tok:
        return None, tok
    sheet, ref = tok.rsplit("!", 1)
    if sheet.startswith("'"):
        sheet = sheet[1:-1].replace("''", "'")
    return sheet, ref


def parse_ref(kind, tok):
    """-> (sheet or None, r1, c1, r2, c2), 0-based."""
    sheet, ref = split_sheet(tok)
    a, _, b = ref.partition(":")
    b = b or a
    if kind == "ref":
        (c1, r1), (c2, r2) = (re.fullmatch(r"\$?([A-Za-z]+)\$?(\d+)", x).groups() for x in (a, b))
        r1, r2, c1, c2 = int(r1) - 1, int(r2) - 1, col_index(c1), col_index(c2)
    elif kind == "cols":
        r1, r2, c1, c2 = 0, MAXR - 1, col_index(a.strip("$")), col_index(b.strip("$"))
    else:
        r1, r2, c1, c2 = int(a.strip("$")) - 1, int(b.strip("$")) - 1, 0, MAXC - 1
    return sheet, min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)


def remap_ref(kind, ref, span_r, span_c):
    """One reference (no sheet prefix) through row/column maps.
    span_r / span_c: (first, last) original 0-based index -> new (first, last), or None if all deleted.
    Returns the new reference text ($ kept) or '#REF!'."""
    parts = [p for p in re.findall(r"(\$?)([A-Za-z]*)(\$?)(\d*)", ref) if p[1] or p[3]]
    if not parts:
        return ref
    rows = [int(p[3]) - 1 for p in parts if p[3]]
    cols = [col_index(p[1]) for p in parts if p[1]]
    nr = span_r(min(rows), max(rows)) if rows else None
    nc = span_c(min(cols), max(cols)) if cols else None
    if (rows and nr is None) or (cols and nc is None):
        return "#REF!"
    out = []
    for i, (cabs, col, rabs, row) in enumerate(parts):
        end = 0 if i == 0 else 1
        if col:
            col = col_letters(nc[end])
        if row:
            row = str(nr[end] + 1)
        out.append(f"{cabs}{col}{rabs}{row}")
    return ":".join(out)


def remap(text, home, target, span_r, span_c):
    """Rewrite every reference to sheet `target` in a formula (references without a sheet
    belong to `home`) after rows/columns moved. See remap_ref."""
    out = []
    for kind, tok in lex(text):
        if kind in ("ref", "cols", "rows"):
            sheet, ref = split_sheet(tok)
            if (sheet or home or "").lower() == target.lower():
                tok = tok[:len(tok) - len(ref)] + remap_ref(kind, ref, span_r, span_c)
        out.append(tok)
    return "".join(out)


def quote_sheet(name):
    """Sheet name as written in a formula: quoted unless it is a plain word."""
    plain = re.fullmatch(r"[A-Za-z_][\w.]*", name) and not re.fullmatch(r"[A-Za-z]{1,3}\d+", name)
    return name if plain else "'" + name.replace("'", "''") + "'"


def rename_sheet(formula, old, new):
    """Point references at sheet `old` to `new`; new=None (sheet deleted) turns them into #REF!."""
    out = []
    for kind, tok in lex(formula):
        if kind in ("ref", "cols", "rows") and "!" in tok:
            sheet, ref = split_sheet(tok)
            if sheet.lower() == old.lower():
                tok = "#REF!" if new is None else f"{quote_sheet(new)}!{ref}"
        out.append(tok)
    return "".join(out)


def transpose_refs(formula, src, dst):
    """Formula copied from cell src to dst with Paste special → Transpose: relative references keep their
    distance, with rows and columns swapped (Excel does the same)."""
    sr, sc = src
    tr, tc = dst

    def cell(m):
        cabs, col, rabs, row = m.groups()
        r, c = int(row) - 1, col_index(col)
        dr, dc = (0 if rabs else r - sr), (0 if cabs else c - sc)
        nr = r if rabs else tr + dc
        nc = c if cabs else tc + dr
        if nr < 0 or nc < 0:
            raise ValueError
        return f"{cabs}{col_letters(nc)}{rabs}{nr + 1}"
    out = []
    for kind, tok in lex(formula[1:]):
        if kind == "ref":
            sheet, ref = split_sheet(tok)
            try:
                ref = re.sub(r"(\$?)([A-Za-z]{1,3})(\$?)(\d+)", cell, ref)
            except ValueError:
                return "=#REF!"
            tok = (tok[:len(tok) - len(split_sheet(tok)[1])] if sheet else "") + ref
        out.append(tok)
    return "=" + "".join(out)


def move_refs(formula, home, target, rect, dr, dc):
    """Cells `rect` (r1, c1, r2, c2) on sheet `target` were moved by (dr, dc): references that lie wholly
    inside the block follow it ($ or not), like Excel's cut & paste / drag. home = sheet the formula is on."""
    if not formula.startswith("="):
        return formula
    r1, c1, r2, c2 = rect
    out, changed = [], False
    for kind, tok in lex(formula[1:]):
        if kind == "ref":
            sheet, ref = split_sheet(tok)
            if (sheet or home).lower() == target.lower():
                try:
                    _, a1, b1, a2, b2 = parse_ref("ref", tok)
                except (AttributeError, ValueError):
                    a1 = None
                if a1 is not None and r1 <= a1 and a2 <= r2 and c1 <= b1 and b2 <= c2:
                    def cell(m):
                        cabs, col, rabs, row = m.groups()
                        return f"{cabs}{col_letters(col_index(col) + dc)}{rabs}{int(row) + dr}"
                    ref = re.sub(r"(\$?)([A-Za-z]{1,3})(\$?)(\d+)", cell, ref)
                    tok = (tok[:len(tok) - len(split_sheet(tok)[1])] if sheet else "") + ref
                    changed = True
        out.append(tok)
    return "=" + "".join(out) if changed else formula


def references(formula):
    """[(sheet or None, r1, c1, r2, c2)] referenced by a formula, in order. Works on half-typed text."""
    if not formula.startswith("="):
        return []
    out = []
    for kind, tok in lex(formula[1:]):
        if kind in ("ref", "cols", "rows"):
            try:
                out.append(parse_ref(kind, tok))
            except (AttributeError, ValueError):
                pass
    return out


def shift(formula, dr, dc):
    """Move relative references by (dr, dc), like Excel does when pasting a formula."""
    def cell(m):
        cabs, col, rabs, row = m.groups()
        c = col_index(col) + (0 if cabs else dc)
        r = int(row) - 1 + (0 if rabs else dr)
        if r < 0 or c < 0:
            raise ValueError
        return f"{cabs}{col_letters(c)}{rabs}{r + 1}"

    def col(m):
        c = col_index(m.group(2)) + (0 if m.group(1) else dc)
        if c < 0:
            raise ValueError
        return m.group(1) + col_letters(c)

    def row(m):
        r = int(m.group(2)) + (0 if m.group(1) else dr)
        if r < 1:
            raise ValueError
        return m.group(1) + str(r)

    out = []
    for kind, tok in lex(formula):
        if kind in ("ref", "cols", "rows"):
            sheet, ref = tok.rsplit("!", 1) if "!" in tok else (None, tok)
            try:
                if kind == "ref":
                    ref = re.sub(r"(\$?)([A-Za-z]{1,3})(\$?)(\d+)", cell, ref)
                elif kind == "cols":
                    ref = re.sub(r"(\$?)([A-Za-z]{1,3})", col, ref)
                else:
                    ref = re.sub(r"(\$?)(\d+)", row, ref)
            except ValueError:
                ref = "#REF!"
            tok = f"{sheet}!{ref}" if sheet else ref
        out.append(tok)
    return "".join(out)


def adjust(formula, home, target, axis, at, n):
    """Fix references after rows/columns were inserted (n > 0) or deleted (n < 0) at index `at`
    on sheet `target`. home = sheet the formula lives on. Deleted single cells become #REF!."""
    def move(i):
        """new index or None if deleted"""
        if i < at:
            return i
        if n > 0:
            return i + n
        return None if i < at - n else i + n

    def span(i1, i2):
        a, b = move(i1), move(i2)
        if n < 0 and a is None and b is None:
            return None
        if a is None:
            a = at
        if b is None:
            b = at - 1
        return (a, b) if a <= b else None

    out = []
    for kind, tok in lex(formula):
        if kind in ("ref", "cols", "rows"):
            sheet, ref = split_sheet(tok)
            if (sheet or home).lower() == target.lower() and \
                    not (kind == "cols" and axis == "row") and not (kind == "rows" and axis == "col"):
                prefix = tok[:len(tok) - len(ref)]
                parts = re.findall(r"(\$?)([A-Za-z]*)(\$?)(\d*)", ref)
                parts = [p for p in parts if p[1] or p[3]]
                idx = [(int(p[3]) - 1 if axis == "row" else col_index(p[1])) for p in parts]
                new = span(idx[0], idx[-1]) if len(idx) == 2 else ((move(idx[0]),) * 2 if move(idx[0]) is not None else None)
                if new is None:
                    tok = prefix + "#REF!"
                else:
                    rebuilt = []
                    for p, v in zip(parts, new):
                        cabs, col, rabs, row = p
                        if axis == "row" and row:
                            row = str(v + 1)
                        elif axis == "col" and col:
                            col = col_letters(v)
                        rebuilt.append(f"{cabs}{col}{rabs}{row}")
                    tok = prefix + ":".join(rebuilt)
        out.append(tok)
    return "".join(out)


def permute(formula, home, target, axis, perm):
    """Rows / columns of sheet `target` were reordered: index i is now perm(i). References follow their
    cells; a range follows its two ends (re-sorted) - but a range that covers every reordered line
    (perm.span) holds the same cells afterwards and stays as it is. home = the formula's sheet."""
    span = getattr(perm, "span", None)
    out = []
    for kind, tok in lex(formula):
        if kind in ("ref", "cols", "rows"):
            sheet, ref = split_sheet(tok)
            if (sheet or home).lower() == target.lower() and \
                    not (kind == "cols" and axis == "row") and not (kind == "rows" and axis == "col"):
                prefix = tok[:len(tok) - len(ref)]
                parts = [p for p in re.findall(r"(\$?)([A-Za-z]*)(\$?)(\d*)", ref) if p[1] or p[3]]
                idx = [(int(p[3]) - 1 if axis == "row" else col_index(p[1])) for p in parts]
                if len(idx) == 2 and span and min(idx) <= span[0] and max(idx) >= span[1]:
                    out.append(tok)                 # =SUM(A2:D2) with B moved after D: same cells, same range
                    continue
                new = [perm(i) for i in idx]
                if len(new) == 2:
                    new = sorted(new)
                rebuilt = []
                for p, v in zip(parts, new):
                    cabs, col, rabs, row = p
                    if axis == "row" and row:
                        row = str(v + 1)
                    elif axis == "col" and col:
                        col = col_letters(v)
                    rebuilt.append(f"{cabs}{col}{rabs}{row}")
                tok = prefix + ":".join(rebuilt)
        out.append(tok)
    return "".join(out)


# ---------- parser ----------

class Parser:
    """Precedence climbing, Excel order: comparison < & < +- < */ < ^ < unary minus < %."""

    def __init__(self, text):
        self.toks = [(k, t) for k, t in lex(text) if k != "ws"]
        self.i = 0

    def peek(self):
        return self.toks[self.i] if self.i < len(self.toks) else (None, None)

    def take(self):
        tok = self.peek()
        self.i += 1
        return tok

    def expect(self, text):
        if self.take()[1] != text:
            raise FormulaError(NAME)

    def parse(self):
        node = self.compare()
        if self.i != len(self.toks):
            raise FormulaError(NAME)
        return node

    def binary(self, ops, sub):
        node = sub()
        while self.peek()[0] == "op" and self.peek()[1] in ops:
            op = self.take()[1]
            node = ("bin", op, node, sub())
        return node

    def compare(self):
        return self.binary(("=", "<>", "<", ">", "<=", ">="), self.concat)

    def concat(self):
        return self.binary(("&",), self.additive)

    def additive(self):
        return self.binary(("+", "-"), self.term)

    def term(self):
        return self.binary(("*", "/"), self.power)

    def power(self):
        return self.binary(("^",), self.unary)

    def unary(self):
        if self.peek() in (("op", "-"), ("op", "+")):
            op = self.take()[1]
            node = self.unary()
            return ("neg", node) if op == "-" else node
        node = self.primary()
        while self.peek() == ("op", "%"):
            self.take()
            node = ("pct", node)
        return node

    def call_args(self):
        self.expect("(")
        args = []
        if self.peek() != ("op", ")"):
            while True:
                if self.peek()[0] == "op" and self.peek()[1] in (",", ";", ")"):
                    args.append(("missing",))
                else:
                    args.append(self.compare())
                if self.peek()[0] == "op" and self.peek()[1] in (",", ";"):
                    self.take()
                    continue
                break
        self.expect(")")
        return args

    def primary(self):
        kind, tok = self.take()
        if kind == "num":
            return ("val", float(tok))
        if kind == "str":
            return ("val", tok[1:-1].replace('""', '"'))
        if kind == "bool":
            return ("val", tok.upper() == "TRUE")
        if kind == "err":
            return ("val", ERRORS.get(tok.upper(), XLError(tok.upper())))
        if kind in ("ref", "cols", "rows"):
            return ("ref",) + parse_ref(kind, tok)
        if kind == "table":      # structured reference: Table1[Col], Table1[#All], [@Col]…
            return ("table", tok)
        if kind == "ext":        # another workbook: '[Book2.xlsx]Sheet1'!A1 or [1]Sheet1!A1
            return ("ext", tok)
        if kind == "spillref":   # A1# = the whole block spilled by A1's formula
            sheet, r1, c1, _, _ = parse_ref("ref", tok[:-1])
            return ("spillref", sheet, r1, c1)
        if kind == "array":      # {1,2;3,4}
            rows = []
            for part in tok[1:-1].split(";"):
                row = []
                for item in part.split(","):
                    sub = Parser(item.strip()).unary() if item.strip() else ("val", None)
                    if sub[0] == "neg":
                        sub = ("val", -sub[1][1])
                    if sub[0] != "val":
                        raise FormulaError(VALUE)
                    row.append(sub[1])
                rows.append(row)
            return ("val", Array(rows))
        if kind == "func":
            name = re.sub(r"^_xl(fn|ws)\.", "", tok, flags=re.I).upper()
            self.expect("(")
            args = []
            if self.peek() != ("op", ")"):
                while True:
                    if self.peek()[0] == "op" and self.peek()[1] in (",", ";", ")"):
                        args.append(("missing",))
                    else:
                        args.append(self.compare())
                    if self.peek()[0] == "op" and self.peek()[1] in (",", ";"):
                        self.take()
                        continue
                    break
            self.expect(")")
            node = ("call", name, args)
            while self.peek() == ("op", "("):     # LAMBDA(x, x*2)(5): call what the expression gives
                node = ("apply", node, self.call_args())
            return node
        if (kind, tok) == ("op", "("):
            node = self.compare()
            self.expect(")")
            return node
        if kind == "name":   # a defined name (named range), looked up when calculating
            node = ("name", tok)
            while self.peek() == ("op", "("):     # a name holding a LAMBDA, called like a function
                node = ("apply", node, self.call_args())
            return node
        raise FormulaError(NAME)


# ---------- values ----------

def serial(d):
    return (d - EPOCH).total_seconds() / 86400


def from_serial(n):
    return EPOCH + dt.timedelta(days=float(n))


def literal(t):
    """Typed value of a plain (non-formula) cell."""
    if t == "":
        return None
    n = parse_number(t)
    if n is not None:
        return n
    u = t.upper()
    if u in ("TRUE", "FALSE"):
        return u == "TRUE"
    if u in ERRORS:
        return ERRORS[u]
    d = parse_date(t)
    if d is not None:
        return Serial(serial(d))
    return t


def num(v):
    if isinstance(v, XLError):
        raise FormulaError(v)
    if v is None or v is MISSING:
        return 0.0
    if isinstance(v, bool):
        return float(v)
    if isinstance(v, (int, float)):
        return v
    if isinstance(v, Range):
        return num(v.scalar())
    if isinstance(v, str):
        s = v.strip()
        n = parse_number(s)
        if n is not None:
            return n
        d = parse_date(s) or parse_date_text(s)       # "15 Oct 2026" + 1 works, like Excel
        if d is not None:
            return Serial(serial(d))
    raise FormulaError(VALUE)


def text(v):
    if isinstance(v, XLError):
        raise FormulaError(v)
    if isinstance(v, Range):
        return text(v.scalar())
    if v is None or v is MISSING:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, Serial):
        d = from_serial(v)
        return d.strftime("%Y-%m-%d") if d.time() == dt.time() else d.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() and abs(v) < 1e15 else f"{v:.15g}"
    if isinstance(v, str):
        return v                       # as it is (an =IMAGE() value keeps its picture for the grid)
    return str(v)


def truthy(v):
    if isinstance(v, Range):
        v = v.scalar()
    if isinstance(v, XLError):
        raise FormulaError(v)
    if isinstance(v, str):
        if v.upper() in ("TRUE", "FALSE"):
            return v.upper() == "TRUE"
        raise FormulaError(VALUE)
    return bool(num(v))


def rank(v):
    return 2 if isinstance(v, bool) else 1 if isinstance(v, str) else 0


def compare(a, b):
    """-> -1 / 0 / 1 with Excel ordering: numbers < text < booleans; text is case-insensitive."""
    for v in (a, b):
        if isinstance(v, XLError):
            raise FormulaError(v)
    if a is None or a is MISSING:
        a = "" if isinstance(b, str) else False if isinstance(b, bool) else 0.0
    if b is None or b is MISSING:
        b = "" if isinstance(a, str) else False if isinstance(a, bool) else 0.0
    ra, rb = rank(a), rank(b)
    if ra != rb:
        return -1 if ra < rb else 1
    if ra == 1:
        a, b = a.casefold(), b.casefold()
    return (a > b) - (a < b)


class Range:
    def __init__(self, eng, sheet, r1, c1, r2, c2):
        s = eng.book[sheet]
        if eng.stack:   # the formula being calculated depends on this whole block (also cells added later)
            eng.range_deps.setdefault(eng.stack[-1], set()).add((sheet, r1, c1, r2, c2))
        if r2 >= MAXR - 1:   # whole column: stop at the data
            r2 = max(len(s.rows) - 1, r1)
        if c2 >= MAXC - 1:   # whole row
            c2 = max(max((len(x) for x in s.rows), default=1) - 1, c1)
        self.eng, self.sheet, self.r1, self.c1, self.r2, self.c2 = eng, sheet, r1, c1, r2, c2
        self.h, self.w = r2 - r1 + 1, c2 - c1 + 1

    def at(self, i, j):
        return self.eng.cell(self.sheet, self.r1 + i, self.c1 + j)

    def flat(self):
        for i in range(self.h):
            for j in range(self.w):
                yield self.at(i, j)

    def scalar(self):
        if self.h == 1 and self.w == 1:
            return self.at(0, 0)
        raise FormulaError(VALUE)


class Array(Range):
    """A calculated block of values (SEQUENCE, FILTER, SORT…): works wherever a range does and
    spills into the cells below / right of its formula."""

    def __init__(self, rows):
        rows = [list(r) for r in rows] or [[]]
        width = max((len(r) for r in rows), default=0)
        self.rows = [r + [NA] * (width - len(r)) for r in rows]
        self.h, self.w = len(self.rows), width
        self.eng = self.sheet = None
        self.r1 = self.c1 = 0
        self.r2, self.c2 = self.h - 1, self.w - 1

    def at(self, i, j):
        return self.rows[i][j]


class Lambda:
    """LAMBDA(params…, body): a formula you can call; names in scope when it was made are kept."""

    def __init__(self, params, body, scope):
        self.params, self.body, self.scope = params, body, scope

    def __repr__(self):
        return "#LAMBDA"


def as_rows(v):
    """Any value / range / array -> list of rows of values."""
    if isinstance(v, Range):
        return [[v.at(i, j) for j in range(v.w)] for i in range(v.h)]
    return [[v]]


def is_block(v):
    """A range / array of more than one value (formulas then work value by value)."""
    return isinstance(v, Range) and (v.h > 1 or v.w > 1)


def pick(rows, i, j):
    """Value (i, j) of a block, a single row / column repeating across the other way (Excel's broadcasting)."""
    if len(rows) == 1:
        i = 0
    if len(rows[0]) == 1:
        j = 0
    return rows[i][j] if i < len(rows) and j < len(rows[i]) else NA


def each(v, f):
    """f applied to every value of a block -> Array (errors stay in their place)."""
    out = []
    for row in as_rows(v):
        line = []
        for x in row:
            try:
                line.append(f(x))
            except FormulaError as e:
                line.append(e.err)
        out.append(line)
    return Array(out)


# one-value functions that work value by value when given a block (Excel's implicit array behaviour)
LIFTED = set("""TEXT LEN LEFT RIGHT MID UPPER LOWER PROPER TRIM CLEAN VALUE DATEVALUE TIMEVALUE NUMBERVALUE YEAR MONTH DAY
HOUR MINUTE SECOND WEEKDAY WEEKNUM ISOWEEKNUM EOMONTH EDATE DATE TIME ROUND ROUNDUP ROUNDDOWN INT TRUNC ABS SQRT
MOD POWER EXP LN LOG LOG10 SIGN CEILING FLOOR CEILING.MATH FLOOR.MATH MROUND EVEN ODD ISNUMBER ISTEXT ISNONTEXT ISBLANK
ISLOGICAL ISERROR ISERR ISNA ISEVEN ISODD N T SEARCH FIND SUBSTITUTE REPLACE REPT EXACT CHAR CODE UNICHAR
UNICODE FIXED DOLLAR NOT IF IFERROR IFNA SIN COS TAN RADIANS DEGREES FACT CONCATENATE TEXTBEFORE
TEXTAFTER""".split())       # (not CONCAT / TEXTJOIN / SUM…: those take whole ranges on purpose)


VOLATILE = re.compile(r"\b(NOW|TODAY|RAND|RANDBETWEEN|RANDARRAY|OFFSET|INDIRECT|CELL|INFO)\s*\(|\[", re.I)
ARRAY_FUNCS = re.compile(r"\b(SEQUENCE|FILTER|UNIQUE|SORT|SORTBY|TRANSPOSE|RANDARRAY|TEXTSPLIT|VSTACK|HSTACK|"
                         r"TOCOL|TOROW|TAKE|DROP|CHOOSECOLS|CHOOSEROWS|WRAPROWS|WRAPCOLS|MMULT|FREQUENCY)\s*\(|"
                         r"^=\s*[^(),\"]*:[^(),\"]*$|[-+*/^&]\s*\$?[A-Za-z]{1,3}\$?\d+:", re.I)


# ---------- engine ----------

class Engine:
    def __init__(self):
        self.book, self.names, self.ids = {}, {}, {}
        self.memo, self.busy, self.asts = {}, set(), {}
        self.too_deep = False
        self.defined = {}   # defined names: lower-case name -> reference text ("Sheet1!$A$1:$B$5")
        self.stack = []     # (sheet, r, c) of the cells being calculated, innermost last
        self.iterate = None         # (max iterations, max change) when circular references may iterate
        self.prev = {}              # last iteration's values of cells in a cycle
        self.cycle = False
        self.array_cells = None     # {(sheet, r, c)} formulas that may spill; None = scan the book again
        self.spill = None           # {(sheet, r, c): value} cells filled by spilling formulas (None = not built)
        self.spill_size = {}        # anchor (sheet, r, c) -> (h, w) of its spilled block
        self.building = False
        self.tables = {}            # lower-case Excel table name -> {"sheet", "range", "columns", "header", "totals"}
        self.externals = None       # fn(reference text) -> value / Array for links to other workbooks
        self.deps = {}              # formula cell -> {cells it read}
        self.range_deps = {}        # formula cell -> {(sheet, r1, c1, r2, c2) blocks it read}
        self.volatile = set()       # formula cells using NOW / RAND / OFFSET / INDIRECT…: always recalculated
        self.env = []               # LET / LAMBDA names in scope: [{lower-case name: value}], innermost last

    def set_iteration(self, props):
        """Iterative calculation (File settings): circular references are calculated over and over."""
        props = props or {}
        self.iterate = (int(props.get("iterate_count", 100)), float(props.get("iterate_delta", 0.001))) \
            if props.get("iterate") else None
        self.clear()

    def set_tables(self, tables):
        self.tables = {t["name"].lower(): t for t in tables}
        self.clear()

    def note(self, sheet, r, c, text):
        """A cell's text changed: keep the list of possibly spilling formulas up to date."""
        if self.array_cells is None:
            return
        name = self.ids.get(id(sheet))
        key = (name, r, c)
        if text and text[0] == "=" and ARRAY_FUNCS.search(text):
            self.array_cells.add(key)
        else:
            self.array_cells.discard(key)

    def rescan(self):
        self.array_cells = None

    def ensure_spill(self):
        """Calculate every possibly-spilling formula so the cells they fill are known."""
        if self.spill is not None or self.building:
            return
        if self.array_cells is None:
            self.array_cells = {(name, r, c) for name, s in self.book.items() for r, row in enumerate(s.rows)
                                for c, t in enumerate(row) if t[:1] == "=" and ARRAY_FUNCS.search(t)}
        self.building = True
        self.spill, self.spill_size = {}, {}
        try:
            for key in sorted(self.array_cells):
                self.cell(*key)
        finally:
            self.building = False

    def spilled(self, sheet_obj, r, c):
        """Text shown in an empty cell that a dynamic-array formula spills into (None if none)."""
        if not self.array_cells and self.array_cells is not None:
            return None
        self.ensure_spill()
        name = self.ids.get(id(sheet_obj))
        v = self.spill.get((name, r, c)) if self.spill else None
        if v is None:
            return None
        return text(v) if not isinstance(v, XLError) else str(v)

    def _spill(self, key, arr):
        """Anchor `key` returned a block: fill the empty cells it covers, or #SPILL! if one isn't empty."""
        sheet, r, c = key
        rows = self.book[sheet].rows
        for i in range(arr.h):
            for j in range(arr.w):
                if (i or j) and r + i < len(rows) and c + j < len(rows[r + i]) and rows[r + i][c + j] != "":
                    return SPILL
        if self.spill is None:
            self.spill = {}
        for i in range(arr.h):
            for j in range(arr.w):
                if i or j:
                    self.spill[(sheet, r + i, c + j)] = arr.at(i, j)
        self.spill_size[key] = (arr.h, arr.w)
        v = arr.at(0, 0)
        return 0.0 if v is None else v

    def set_names(self, names):
        self.defined = {n.lower(): t for n, t in (names or {}).items()}
        self.clear()

    def set_book(self, book):
        self.book = book
        self.names = {n.lower(): n for n in book}
        self.ids = {id(s): n for n, s in book.items()}
        self.array_cells = None
        self.clear()

    def clear(self):
        self.memo = {}
        self.spill = None
        self.deps, self.range_deps, self.volatile = {}, {}, set()
        self.loops = {}             # formula cell -> {cells it read that were still being calculated: #CIRC!}

    def invalidate(self, cells):
        """Some cells' texts changed: forget the results of exactly those cells and every formula that
        (directly or through others) reads them. Falls back to recalculating everything when that is
        simpler (spilling formulas, iterative calculation)."""
        if self.iterate or self.array_cells or self.too_deep:
            self.clear()
            return None
        rev = {}
        for f, ps in self.deps.items():
            for p in ps:
                rev.setdefault(p, set()).add(f)
        by_sheet = {}
        for f, blocks in self.range_deps.items():
            for b in blocks:
                by_sheet.setdefault(b[0], []).append((b, f))
        todo = list(cells) + list(self.volatile)
        seen = set(todo)
        while todo:
            k = todo.pop()
            nxt = set(rev.get(k, ()))
            sheet, r, c = k
            for (_, r1, c1, r2, c2), f in by_sheet.get(sheet, ()):
                if r1 <= r <= r2 and c1 <= c <= c2:
                    nxt.add(f)
            for f in nxt - seen:
                seen.add(f)
                todo.append(f)
        for k in seen:
            self.memo.pop(k, None)
            self.deps.pop(k, None)
            self.range_deps.pop(k, None)
            self.loops.pop(k, None)
            self.volatile.discard(k)
        return seen

    def display(self, sheet_obj, r, c):
        """Text to show for a formula cell."""
        name = self.ids.get(id(sheet_obj))
        if name is None:
            return ""
        self.too_deep = False
        self.cycle = False
        v = self.cell(name, r, c)
        if self.cycle and self.iterate:     # circular: iterate until the values settle
            count, delta = self.iterate
            for _ in range(count):
                cur = {k: x for k, x in self.memo.items() if isinstance(x, (int, float)) and not isinstance(x, bool)}
                change = max((abs(x - self.prev.get(k, 0.0)) for k, x in cur.items()), default=0.0)
                self.prev.update(cur)
                if change < delta:
                    break
                self.memo = {}
                v = self.cell(name, r, c)
        if self.too_deep:   # long chain (e.g. running totals): calculate top-down, then retry
            self.memo = {}
            for rr in range(r + 1):
                for cc in range(len(sheet_obj.rows[rr]) if rr < len(sheet_obj.rows) else 0):
                    self.cell(name, rr, cc)
            v = self.cell(name, r, c)
        cached = sheet_obj.cache.get((r, c))
        if v == NAME and cached is not None:   # a function we don't know: keep Excel's last result
            return cached
        return text(v) if not isinstance(v, XLError) else str(v)

    def circle(self, sheet_obj, r, c, limit=20000):
        """Why a cell shows #CIRC!: the loop of formula cells it runs into -> [(sheet, r, c), ..., the first one
        again] (just [x, x] when a cell uses itself), or None. The loop may not start at this cell (=SUM(B3:E3)
        when B3 uses itself)."""
        name = self.ids.get(id(sheet_obj))
        if name is None:
            return None
        self.display(sheet_obj, r, c)                    # (calculating it records what each formula read)

        def is_formula(k):
            rows = self.book[k[0]].rows
            t = rows[k[1]][k[2]] if k[1] < len(rows) and k[2] < len(rows[k[1]]) else ""
            return len(t) > 1 and t[0] == "="

        def reads(k):
            found = set(self.deps.get(k, ())) | self.loops.get(k, set())
            budget = limit
            for sheet, r1, c1, r2, c2 in self.range_deps.get(k, ()):
                rows = self.book[sheet].rows
                for rr in range(r1, min(r2, len(rows) - 1) + 1):
                    for cc in range(c1, min(c2, len(rows[rr]) - 1) + 1):
                        budget -= 1
                        if budget < 0:
                            return found
                        found.add((sheet, rr, cc))
            return {x for x in found if is_formula(x)}

        start = (name, r, c)
        path, on_path, done = [start], {start}, set()
        todo = [iter(sorted(reads(start)))]
        while todo:                                      # depth first, remembering the way here
            nxt = next(todo[-1], None)
            if nxt is None:                              # all it reads looked at: step back
                todo.pop()
                left = path.pop()
                done.add(left)
                on_path.discard(left)
                continue
            if nxt in on_path:                           # back to a cell on the way: that's the loop
                return path[path.index(nxt):] + [nxt]
            if nxt in done or len(path) > 200:
                continue
            path.append(nxt)
            on_path.add(nxt)
            todo.append(iter(sorted(reads(nxt))))
        return None

    def cell(self, sheet, r, c):
        key = (sheet, r, c)
        if self.stack and self.stack[-1] != key:      # the formula being calculated reads this cell
            self.deps.setdefault(self.stack[-1], set()).add(key)
        if key in self.memo:
            return self.memo[key]
        rows = self.book[sheet].rows
        t = rows[r][c] if r < len(rows) and c < len(rows[r]) else ""
        if len(t) > 1 and t[0] == "=":
            if key in self.busy:
                if self.iterate:
                    self.cycle = True
                    return self.prev.get(key, 0.0)
                if self.stack:                            # (who read it: the loop's last step, for circle())
                    self.loops.setdefault(self.stack[-1], set()).add(key)
                return CIRC
            self.busy.add(key)
            self.deps.pop(key, None)
            self.range_deps.pop(key, None)
            self.loops.pop(key, None)
            if VOLATILE.search(t):
                self.volatile.add(key)
            self.stack.append(key)    # the cell being calculated (ROW() / COLUMN() without arguments)
            try:
                v = self.evaluate(t, sheet, spill=True)
                if isinstance(v, Range):
                    v = self._spill(key, v)
            finally:
                self.busy.discard(key)
                self.stack.pop()
        elif t == "" and self.array_cells != set():
            self.ensure_spill()
            v = self.spill.get(key) if self.spill else None
        else:
            v = literal(t)
        self.memo[key] = v
        return v

    def evaluate(self, formula, sheet, spill=False):
        try:
            ast = self.asts.get(formula)
            if ast is None:
                if len(self.asts) > 20000:
                    self.asts.clear()
                ast = self.asts[formula] = Parser(formula[1:]).parse()
            v = self.ev(ast, sheet)
            if isinstance(v, Range):
                if spill and (v.h > 1 or v.w > 1):
                    return v                    # a block: the cell spills it (Engine.cell)
                v = v.scalar()
            if isinstance(v, Lambda):        # a LAMBDA that wasn't called
                return CALC
            if v is None or v is MISSING:
                return 0.0
            if isinstance(v, float) and not math.isfinite(v):
                return NUM
            return v
        except FormulaError as e:
            return e.err
        except RecursionError:
            self.too_deep = True
            return NUM
        except (ValueError, TypeError, OverflowError, ZeroDivisionError, IndexError):
            return VALUE

    def sheet_name(self, name, default):
        if name is None:
            return default
        real = self.names.get(name.lower())
        if real is None:
            raise FormulaError(REF)
        return real

    def ev(self, node, sheet):
        kind = node[0]
        if kind == "val":
            return node[1]
        if kind == "missing":
            return MISSING
        if kind == "table":
            return self.table_ref(node[1], sheet)
        if kind == "ext":
            if self.externals is None:
                raise FormulaError(REF)
            return self.externals(node[1])
        if kind == "spillref":
            _, name, r, c = node
            name = self.sheet_name(name, sheet)
            self.cell(name, r, c)
            h, w = self.spill_size.get((name, r, c), (1, 1))
            return Range(self, name, r, c, r + h - 1, c + w - 1)
        if kind == "apply":                       # (expression)(args): a LAMBDA called on the spot
            return self.call_lambda(self.ev(node[1], sheet), node[2], sheet)
        if kind == "name":
            low = node[1].lower()
            for scope in reversed(self.env):      # LET / LAMBDA names first
                if low in scope:
                    return scope[low]
            text = self.defined.get(low)
            if text is None and node[1].lower() in self.tables:   # a table's name = its data
                return self.table_ref(node[1] + "[]", sheet)
            if text is None:
                raise FormulaError(NAME)
            target = self.asts.get("name:" + text)
            if target is None:
                target = self.asts["name:" + text] = Parser(text.lstrip("=")).parse()
            return self.ev(target, sheet)
        if kind == "ref":
            _, name, r1, c1, r2, c2 = node
            name = self.sheet_name(name, sheet)
            if r1 == r2 and c1 == c2:
                return self.cell(name, r1, c1)
            return Range(self, name, r1, c1, r2, c2)
        if kind == "neg":
            v = self.ev(node[1], sheet)
            if is_block(v):                       # -range / --(A1:A9>5): each value
                return each(v, lambda x: -num(x))
            return -num(v)
        if kind == "pct":
            v = self.ev(node[1], sheet)
            if is_block(v):
                return each(v, lambda x: num(x) / 100)
            return num(v) / 100
        if kind == "bin":
            return self.binop(node[1], self.ev(node[2], sheet), self.ev(node[3], sheet))
        if kind == "call":
            f = FUNCS.get(node[1])
            if f is None:                         # a name holding a LAMBDA (LET name or defined name)
                low = node[1].lower()
                target = next((s[low] for s in reversed(self.env) if low in s), None)
                if target is None and low in self.defined:
                    target = self.ev(("name", node[1]), sheet)
                if isinstance(target, Lambda):
                    return self.call_lambda(target, node[2], sheet)
                raise FormulaError(NAME)
            func, lazy = f
            if lazy == "raw":   # gets the argument expressions themselves (ROW, COLUMN, OFFSET)
                return func(self, node[2], sheet)
            if node[1] in LIFTED and node[2]:     # a one-value function given a block: once per value
                return self.lift_call(func, lazy, node[2], sheet)
            if lazy:
                return func(self, [(lambda n=n: self.ev(n, sheet)) for n in node[2]])
            return func(self, [self.ev(n, sheet) for n in node[2]])
        raise FormulaError(NAME)

    def lift_call(self, func, lazy, nodes, sheet):
        """Excel's array rule for one-value functions: TEXT(A2:A500, "yyyy-mm"), LEN(B:B), IF(A1:A9>5, …) give
        a block with the function applied to each value (SUMPRODUCT / SUM / FILTER… then use it)."""
        try:
            first, first_err = self.ev(nodes[0], sheet), None
        except FormulaError as e:
            first, first_err = None, e
        if lazy:                                  # IF / IFERROR…: only the first argument decides
            if first_err is None and is_block(first):
                return self.lift(func, True, [first] + [self.ev_value(n, sheet) for n in nodes[1:]])

            def get_first():
                if first_err is not None:
                    raise first_err
                return first
            return func(self, [get_first] + [(lambda n=n: self.ev(n, sheet)) for n in nodes[1:]])
        if first_err is not None:
            raise first_err
        vals = [first] + [self.ev(n, sheet) for n in nodes[1:]]
        if any(is_block(v) for v in vals):
            return self.lift(func, False, vals)
        return func(self, vals)

    def ev_value(self, node, sheet):
        try:
            return self.ev(node, sheet)
        except FormulaError as e:
            return e.err

    def lift(self, func, lazy, vals):
        blocks = [as_rows(v) if is_block(v) else None for v in vals]
        h = max(len(b) for b in blocks if b)
        w = max(len(b[0]) for b in blocks if b)
        out = []
        for i in range(h):
            row = []
            for j in range(w):
                args = [pick(b, i, j) if b else v for b, v in zip(blocks, vals)]
                try:
                    r = func(self, [(lambda x=x: x) for x in args]) if lazy else func(self, args)
                    if isinstance(r, Range):
                        r = r.at(0, 0) if r.h == 1 and r.w == 1 else VALUE
                except FormulaError as e:
                    r = e.err
                row.append(r)
            out.append(row)
        return Array(out)

    def call_lambda(self, fn, arg_nodes, sheet, values=None):
        """Run a LAMBDA with argument expressions (or already calculated `values`)."""
        if not isinstance(fn, Lambda):
            raise FormulaError(VALUE)
        args = values if values is not None else [self.ev(n, sheet) for n in arg_nodes]
        if len(args) != len(fn.params):
            raise FormulaError(VALUE)
        scope = dict(fn.scope)
        scope.update(zip(fn.params, args))
        self.env.append(scope)
        try:
            return self.ev(fn.body, sheet)
        finally:
            self.env.pop()

    def table_ref(self, tok, sheet):
        """Structured reference -> Range. Tables come from the file (Sheet.extra['tables'])."""
        name, _, inner = tok.partition("[")
        inner = "[" + inner
        if name:
            t = self.tables.get(name.lower())
        else:                                     # [@Col] inside a table: the table holding this cell
            here = self.stack[-1] if self.stack else None
            t = next((t for t in self.tables.values() if here and t["sheet"] == here[0]
                      and t["range"][0] <= here[1] <= t["range"][2] and t["range"][1] <= here[2] <= t["range"][3]), None)
        if t is None:
            raise FormulaError(REF)
        r1, c1, r2, c2 = t["range"]
        head = r1 + 1 if t["header"] else r1          # first data row
        tail = r2 - 1 if t["totals"] else r2          # last data row
        body = inner[1:-1].strip()
        items = re.findall(r"\[([^\[\]]*)\]", body) if body.startswith("[") else ([body] if body else [])
        rows, cols, this_row = (head, tail), (c1, c2), False
        col_names = [n.lower() for n in t["columns"]]
        picked = []
        for it in items:
            it = it.strip()
            low = it.lower()
            if low == "#all":
                rows = (r1, r2)
            elif low == "#data":
                rows = (head, tail)
            elif low == "#headers":
                rows = (r1, r1)
            elif low == "#totals":
                rows = (r2, r2)
            elif low == "#this row" or it.startswith("@"):
                this_row = True
                it = it[1:].strip()
                if it:
                    picked.append(it)
            elif it:
                picked.append(it)
        if picked:
            idx = []
            for p in picked:
                p = p.replace("'", "") if "'" in p else p
                if p.lower() not in col_names:
                    raise FormulaError(REF)
                idx.append(c1 + col_names.index(p.lower()))
            cols = (min(idx), max(idx))
        if this_row:
            if not self.stack:
                raise FormulaError(VALUE)
            r = self.stack[-1][1]
            rows = (r, r)
        if cols[0] == cols[1] and rows[0] == rows[1]:
            return self.cell(t["sheet"], rows[0], cols[0])
        return Range(self, t["sheet"], rows[0], cols[0], rows[1], cols[1])

    def binop(self, op, a, b):
        big = lambda v: isinstance(v, Range) and (v.h > 1 or v.w > 1)   # noqa: E731
        if big(a) or big(b):              # a block with a value / another block: element by element
            ra, rb = as_rows(a), as_rows(b)
            h, w = max(len(ra), len(rb)), max(len(ra[0]), len(rb[0]))
            out = []
            for i in range(h):
                row = []
                for j in range(w):
                    try:
                        row.append(self.binop(op, pick(ra, i, j), pick(rb, i, j)))
                    except FormulaError as e:
                        row.append(e.err)
                out.append(row)
            return Array(out)
        if isinstance(a, XLError):
            raise FormulaError(a)
        if isinstance(b, XLError):
            raise FormulaError(b)
        if isinstance(a, Range):
            a = a.scalar()
        if isinstance(b, Range):
            b = b.scalar()
        if op == "&":
            return text(a) + text(b)
        if op in ("=", "<>", "<", ">", "<=", ">="):
            c = compare(a, b)
            return {"=": c == 0, "<>": c != 0, "<": c < 0, ">": c > 0, "<=": c <= 0, ">=": c >= 0}[op]
        x, y = num(a), num(b)
        if op == "+":
            r = x + y
            return Serial(r) if isinstance(x, Serial) != isinstance(y, Serial) else r
        if op == "-":
            r = x - y
            return Serial(r) if isinstance(x, Serial) and not isinstance(y, Serial) else float(r)
        if op == "*":
            return x * y
        if op == "/":
            if y == 0:
                raise FormulaError(DIV0)
            return x / y
        if op == "^":
            if x == 0 and y < 0:
                raise FormulaError(DIV0)
            try:
                r = math.pow(x, y)
            except (ValueError, OverflowError):
                raise FormulaError(NUM) from None
            return r
        raise FormulaError(NAME)


# ---------- functions ----------

FUNCS = {}


def fn(*names, lazy=False):
    def deco(f):
        for n in names:
            FUNCS[n] = (f, lazy)
        return f
    return deco


def arg(args, i, default=MISSING):
    v = args[i] if i < len(args) else MISSING
    return default if v is MISSING else v


def values(args):
    """Flatten ranges and plain args."""
    for a in args:
        if isinstance(a, Range):
            yield from a.flat()
        elif a is not MISSING:
            yield a


def nums(args):
    """Numbers like SUM sees them: ranges give only numeric cells, direct args are converted."""
    out = []
    for a in args:
        if isinstance(a, Range):
            for v in a.flat():
                if isinstance(v, XLError):
                    raise FormulaError(v)
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    out.append(v)
        elif a is not MISSING:
            out.append(num(a))
    return out


def rnd(x, digits, mode):
    q = Decimal(1).scaleb(-int(digits))
    return float(Decimal(repr(float(x))).quantize(q, rounding=mode))


def as_range(v):
    if not isinstance(v, Range):
        raise FormulaError(VALUE)
    return v


@fn("SUM")
def _sum(e, a): return float(sum(nums(a)))


@fn("AVERAGE")
def _avg(e, a):
    n = nums(a)
    if not n:
        raise FormulaError(DIV0)
    return sum(n) / len(n)


@fn("MIN")
def _min(e, a): return float(min(nums(a), default=0))


@fn("MAX")
def _max(e, a): return float(max(nums(a), default=0))


@fn("MEDIAN")
def _median(e, a):
    n = nums(a)
    if not n:
        raise FormulaError(NUM)
    return float(statistics.median(n))


@fn("PRODUCT")
def _product(e, a): return float(math.prod(nums(a)))


@fn("COUNT")
def _count(e, a):
    n = 0
    for x in a:
        if isinstance(x, Range):
            n += sum(1 for v in x.flat() if isinstance(v, (int, float)) and not isinstance(v, bool))
        elif x is not MISSING:
            try:
                num(x)
                n += 1
            except FormulaError:
                pass
    return float(n)


@fn("COUNTA")
def _counta(e, a): return float(sum(1 for v in values(a) if v is not None and v != ""))


@fn("COUNTBLANK")
def _countblank(e, a): return float(sum(1 for v in as_range(a[0]).flat() if v is None or v == ""))


def _kth(largest):
    def f(e, a):
        n, k = sorted(nums([a[0]]), reverse=largest), int(num(a[1]))
        if not 1 <= k <= len(n):
            raise FormulaError(NUM)
        return float(n[k - 1])
    return f


FUNCS["LARGE"], FUNCS["SMALL"] = (_kth(True), False), (_kth(False), False)


@fn("ROUND")
def _round(e, a): return rnd(num(a[0]), num(arg(a, 1, 0)), ROUND_HALF_UP)


@fn("ROUNDUP")
def _roundup(e, a): return rnd(num(a[0]), num(arg(a, 1, 0)), ROUND_UP)


@fn("ROUNDDOWN", "TRUNC")
def _rounddown(e, a): return rnd(num(a[0]), num(arg(a, 1, 0)), ROUND_DOWN)


@fn("INT")
def _int(e, a): return float(math.floor(num(a[0])))


@fn("ABS")
def _abs(e, a): return abs(num(a[0]))


@fn("SIGN")
def _sign(e, a):
    x = num(a[0])
    return float((x > 0) - (x < 0))


@fn("MOD")
def _mod(e, a):
    x, y = num(a[0]), num(a[1])
    if y == 0:
        raise FormulaError(DIV0)
    return x % y


@fn("POWER")
def _power(e, a): return e.binop("^", a[0], a[1])


@fn("SQRT")
def _sqrt(e, a):
    x = num(a[0])
    if x < 0:
        raise FormulaError(NUM)
    return math.sqrt(x)


@fn("EXP")
def _exp(e, a): return math.exp(num(a[0]))


@fn("LN")
def _ln(e, a):
    x = num(a[0])
    if x <= 0:
        raise FormulaError(NUM)
    return math.log(x)


@fn("LOG")
def _log(e, a):
    x, b = num(a[0]), num(arg(a, 1, 10))
    if x <= 0 or b <= 0 or b == 1:
        raise FormulaError(NUM)
    return math.log(x, b)


@fn("LOG10")
def _log10(e, a): return _log(e, [a[0], 10.0])


@fn("PI")
def _pi(e, a): return math.pi


@fn("RAND")
def _rand(e, a): return random.random()


@fn("RANDBETWEEN")
def _randbetween(e, a): return float(random.randint(math.ceil(num(a[0])), math.floor(num(a[1]))))


# logic

@fn("IF", lazy=True)
def _if(e, a):
    if truthy(a[0]()):
        v = a[1]() if len(a) > 1 else True
    else:
        v = a[2]() if len(a) > 2 else False
    return 0.0 if v is MISSING else v


@fn("IFERROR", lazy=True)
def _iferror(e, a):
    try:
        v = a[0]()
        if isinstance(v, Range):
            v = v.scalar()
        if not isinstance(v, XLError):
            return v
    except FormulaError:
        pass
    return a[1]()


@fn("IFNA", lazy=True)
def _ifna(e, a):
    try:
        v = a[0]()
        if isinstance(v, Range):
            v = v.scalar()
        if v != NA:
            return v
    except FormulaError as err:
        if err.err != NA:
            raise
    return a[1]()


@fn("IFS", lazy=True)
def _ifs(e, a):
    for i in range(0, len(a) - 1, 2):
        if truthy(a[i]()):
            return a[i + 1]()
    raise FormulaError(NA)


def bools(args):
    out = []
    for v in values(args):
        if isinstance(v, XLError):
            raise FormulaError(v)
        if isinstance(v, bool) or isinstance(v, (int, float)):
            out.append(bool(v))
        elif isinstance(v, str) and v.upper() in ("TRUE", "FALSE"):
            out.append(v.upper() == "TRUE")
    if not out:
        raise FormulaError(VALUE)
    return out


@fn("AND")
def _and(e, a): return all(bools(a))


@fn("OR")
def _or(e, a): return any(bools(a))


@fn("XOR")
def _xor(e, a): return sum(bools(a)) % 2 == 1


@fn("NOT")
def _not(e, a): return not truthy(a[0])


@fn("TRUE")
def _true(e, a): return True


@fn("FALSE")
def _false(e, a): return False


@fn("CHOOSE")
def _choose(e, a):
    i = int(num(a[0]))
    if not 1 <= i < len(a):
        raise FormulaError(VALUE)
    return a[i]


# text

@fn("CONCATENATE", "CONCAT")
def _concat(e, a): return "".join(text(v) for v in values(a))


@fn("TEXTJOIN")
def _textjoin(e, a):
    sep, skip = text(a[0]), truthy(a[1])
    parts = [text(v) for v in values(a[2:])]
    return sep.join(p for p in parts if p or not skip)


@fn("LEN")
def _len(e, a): return float(len(text(a[0])))


@fn("UPPER")
def _upper(e, a): return text(a[0]).upper()


@fn("LOWER")
def _lower(e, a): return text(a[0]).lower()


@fn("PROPER")
def _proper(e, a): return text(a[0]).title()


@fn("TRIM")
def _trim(e, a): return re.sub(r" +", " ", text(a[0]).strip())


@fn("LEFT")
def _left(e, a): return text(a[0])[:max(int(num(arg(a, 1, 1))), 0)]


@fn("RIGHT")
def _right(e, a):
    n = max(int(num(arg(a, 1, 1))), 0)
    return text(a[0])[-n:] if n else ""


@fn("MID")
def _mid(e, a):
    s, start, n = text(a[0]), int(num(a[1])), int(num(a[2]))
    if start < 1 or n < 0:
        raise FormulaError(VALUE)
    return s[start - 1:start - 1 + n]


@fn("SUBSTITUTE")
def _substitute(e, a):
    s, old, new = text(a[0]), text(a[1]), text(a[2])
    if len(a) < 4 or a[3] is MISSING:
        return s.replace(old, new) if old else s
    k, pos = int(num(a[3])), -1
    for _ in range(k):
        pos = s.find(old, pos + 1)
        if pos < 0:
            return s
    return s[:pos] + new + s[pos + len(old):]


@fn("REPLACE")
def _replace(e, a):
    s, start, n, new = text(a[0]), int(num(a[1])), int(num(a[2])), text(a[3])
    return s[:start - 1] + new + s[start - 1 + n:]


def _find(case):
    def f(e, a):
        what, s, start = text(a[0]), text(a[1]), int(num(arg(a, 2, 1)))
        pos = (s if case else s.lower()).find(what if case else what.lower(), start - 1)
        if pos < 0:
            raise FormulaError(VALUE)
        return float(pos + 1)
    return f


FUNCS["FIND"], FUNCS["SEARCH"] = (_find(True), False), (_find(False), False)


@fn("REPT")
def _rept(e, a): return text(a[0]) * max(int(num(a[1])), 0)


@fn("EXACT")
def _exact(e, a): return text(a[0]) == text(a[1])


@fn("VALUE")
def _value(e, a): return float(num(text(a[0])))


@fn("TEXT")
def _text(e, a):
    v, f = a[0], text(a[1])
    if isinstance(v, Range):
        v = v.scalar()
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return fmt_number(float(v), f)
    return render(text(v), f)


# dates

@fn("TODAY")
def _today(e, a): return Serial(math.floor(serial(dt.datetime.now())))


@fn("NOW")
def _now(e, a): return Serial(serial(dt.datetime.now()))


@fn("DATE")
def _date(e, a):
    y, m, d = int(num(a[0])), int(num(a[1])), int(num(a[2]))
    if y < 1900:
        y += 1900
    y, m = y + (m - 1) // 12, (m - 1) % 12 + 1
    return Serial(serial(dt.datetime(y, m, 1) + dt.timedelta(days=d - 1)))


@fn("YEAR")
def _year(e, a): return float(from_serial(num(a[0])).year)


@fn("MONTH")
def _month(e, a): return float(from_serial(num(a[0])).month)


@fn("DAY")
def _day(e, a): return float(from_serial(num(a[0])).day)


@fn("WEEKDAY")
def _weekday(e, a): return float((from_serial(num(a[0])).weekday() + 1) % 7 + 1)   # Sunday = 1


@fn("DAYS")
def _days(e, a): return float(num(a[0]) - num(a[1]))


# info

@fn("ISBLANK")
def _isblank(e, a): return (a[0].scalar() if isinstance(a[0], Range) else a[0]) is None


@fn("ISNUMBER")
def _isnumber(e, a):
    v = a[0].scalar() if isinstance(a[0], Range) else a[0]
    return isinstance(v, (int, float)) and not isinstance(v, bool)


@fn("ISTEXT")
def _istext(e, a):
    v = a[0].scalar() if isinstance(a[0], Range) else a[0]
    return isinstance(v, str) and not isinstance(v, XLError)


@fn("N")
def _n(e, a):
    """N(value): a number as it is, TRUE 1, a date its serial, anything else 0 (errors stay errors)."""
    v = a[0].scalar() if isinstance(a[0], Range) else a[0]
    if isinstance(v, XLError):
        raise FormulaError(v)
    if isinstance(v, bool):
        return float(v)
    return v if isinstance(v, (int, float)) else 0.0


@fn("T")
def _t(e, a):
    """T(value): text as it is, anything else empty text."""
    v = a[0].scalar() if isinstance(a[0], Range) else a[0]
    if isinstance(v, XLError):
        raise FormulaError(v)
    return v if isinstance(v, str) else ""


@fn("ISERROR", lazy=True)
def _iserror(e, a):
    try:
        v = a[0]()
        return isinstance(v.scalar() if isinstance(v, Range) else v, XLError)
    except FormulaError:
        return True


@fn("ISNA", lazy=True)
def _isna(e, a):
    try:
        v = a[0]()
        return (v.scalar() if isinstance(v, Range) else v) == NA
    except FormulaError as err:
        return err.err == NA


@fn("NA")
def _na(e, a): raise FormulaError(NA)


@fn("ROWS")
def _rows(e, a): return float(as_range(a[0]).h) if isinstance(a[0], Range) else 1.0


@fn("COLUMNS")
def _columns(e, a): return float(as_range(a[0]).w) if isinstance(a[0], Range) else 1.0


# conditional aggregates

def criteria(c):
    """Excel criteria ("">5"", ""<>x"", ""a*"", 10) -> predicate."""
    if isinstance(c, Range):
        c = c.scalar()
    if not isinstance(c, str) or isinstance(c, XLError):   # a number / boolean: equality
        return lambda v: v is not None and not isinstance(v, XLError) and compare(v, c) == 0
    m = re.match(r"(<=|>=|<>|<|>|=)?(.*)", c, re.S)
    op, rhs = m.group(1) or "=", m.group(2)
    n = parse_number(rhs)
    target = n if n is not None else rhs
    if isinstance(target, str) and op in ("=", "<>"):
        if target == "":
            empty = lambda v: v is None or v == ""
            return empty if op == "=" else (lambda v: not empty(v))
        pat = re.compile(re.escape(target).replace(r"\*", ".*").replace(r"\?", "."), re.I | re.S)
        hit = lambda v: isinstance(v, str) and not isinstance(v, XLError) and pat.fullmatch(v) is not None
        return hit if op == "=" else (lambda v: not hit(v))

    def test(v):
        if v is None or isinstance(v, XLError) or rank(v) != rank(target):
            return op == "<>"
        k = compare(v, target)
        return {"=": k == 0, "<>": k != 0, "<": k < 0, ">": k > 0, "<=": k <= 0, ">=": k >= 0}[op]
    return test


def matching(pairs):
    """pairs: [(range, criteria)] -> list of (i, j) positions where all criteria hold."""
    first = as_range(pairs[0][0])
    tests = [(as_range(r), criteria(c)) for r, c in pairs]
    out = []
    for i in range(first.h):
        for j in range(first.w):
            if all(t(r.at(i, j)) if i < r.h and j < r.w else False for r, t in tests):
                out.append((i, j))
    return out


def picked(rng, pos):
    vals = [rng.at(i, j) for i, j in pos if i < rng.h and j < rng.w]
    return [v for v in vals if isinstance(v, (int, float)) and not isinstance(v, bool)]


@fn("SUMIF")
def _sumif(e, a):
    rng = as_range(a[0])
    total = as_range(arg(a, 2, rng))
    return float(sum(picked(total, matching([(rng, a[1])]))))


@fn("COUNTIF")
def _countif(e, a): return float(len(matching([(a[0], a[1])])))


@fn("AVERAGEIF")
def _averageif(e, a):
    rng = as_range(a[0])
    v = picked(as_range(arg(a, 2, rng)), matching([(rng, a[1])]))
    if not v:
        raise FormulaError(DIV0)
    return sum(v) / len(v)


def pairs_of(a):
    return [(a[i], a[i + 1]) for i in range(0, len(a) - 1, 2)]


@fn("SUMIFS")
def _sumifs(e, a): return float(sum(picked(as_range(a[0]), matching(pairs_of(a[1:])))))


@fn("COUNTIFS")
def _countifs(e, a): return float(len(matching(pairs_of(a))))


@fn("AVERAGEIFS")
def _averageifs(e, a):
    v = picked(as_range(a[0]), matching(pairs_of(a[1:])))
    if not v:
        raise FormulaError(DIV0)
    return sum(v) / len(v)


@fn("MAXIFS")
def _maxifs(e, a): return float(max(picked(as_range(a[0]), matching(pairs_of(a[1:]))), default=0))


@fn("MINIFS")
def _minifs(e, a): return float(min(picked(as_range(a[0]), matching(pairs_of(a[1:]))), default=0))


# lookup

def scalar(v):
    return v.scalar() if isinstance(v, Range) else v


def find_exact(look, seq):
    pred = criteria(look) if isinstance(look, str) and ("*" in look or "?" in look) else None
    for k, v in enumerate(seq):
        if pred(v) if pred else (v is not None and rank(v) == rank(look) and compare(v, look) == 0):
            return k
    return None


def find_sorted(look, seq, descending=False):
    """Largest value <= look (or smallest >= look when descending), like approximate MATCH."""
    best = None
    for k, v in enumerate(seq):
        if v is None or isinstance(v, XLError) or rank(v) != rank(look):
            continue
        c = compare(v, look)
        if (c <= 0) if not descending else (c >= 0):
            best = k
        elif not descending:
            break
    return best


def _lookup(vertical):
    def f(e, a):
        look, rng, idx = scalar(a[0]), as_range(a[1]), int(num(a[2]))
        approx = truthy(arg(a, 3, True))
        size = rng.w if vertical else rng.h
        if not 1 <= idx <= size:
            raise FormulaError(REF)
        n = rng.h if vertical else rng.w
        keys = [rng.at(k, 0) if vertical else rng.at(0, k) for k in range(n)]
        k = find_sorted(look, keys) if approx else find_exact(look, keys)
        if k is None:
            raise FormulaError(NA)
        return rng.at(k, idx - 1) if vertical else rng.at(idx - 1, k)
    return f


FUNCS["VLOOKUP"], FUNCS["HLOOKUP"] = (_lookup(True), False), (_lookup(False), False)


@fn("MATCH")
def _match(e, a):
    look, rng, kind = scalar(a[0]), as_range(a[1]), int(num(arg(a, 2, 1)))
    seq = list(rng.flat())
    k = find_exact(look, seq) if kind == 0 else find_sorted(look, seq, descending=kind < 0)
    if k is None:
        raise FormulaError(NA)
    return float(k + 1)


@fn("INDEX")
def _index(e, a):
    rng = as_range(a[0])
    r, c = int(num(arg(a, 1, 0))), int(num(arg(a, 2, 0)))
    if rng.h == 1 and len(a) < 3:   # one-row range: the single index is a column
        r, c = 1, r
    if r == 0 and c == 0:
        raise FormulaError(VALUE)
    r, c = r or 1, c or 1
    if not (1 <= r <= rng.h and 1 <= c <= rng.w):
        raise FormulaError(REF)
    return rng.at(r - 1, c - 1)


@fn("XLOOKUP")
def _xlookup(e, a):
    look, keys, out = scalar(a[0]), as_range(a[1]), as_range(a[2])
    k = find_exact(look, list(keys.flat()))
    if k is None:
        if len(a) > 3 and a[3] is not MISSING:
            return a[3]
        raise FormulaError(NA)
    return out.at(k, 0) if keys.w == 1 else out.at(0, k)   # vertical or horizontal lookup


# more functions (finance, statistics, dates, references...) register themselves into FUNCS
import functions_array  # noqa: E402,F401
import functions_extra  # noqa: E402,F401
import functions_lambda  # noqa: E402,F401
