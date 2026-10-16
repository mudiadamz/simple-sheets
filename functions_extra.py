"""More spreadsheet functions: finance, statistics, math, dates/times, text and references.
Registers into formula.FUNCS (imported at the end of formula.py)."""
import calendar
import math
import statistics
from collections import Counter

from formula import (
    DIV0,
    NA,
    NUM,
    REF,
    VALUE,
    FormulaError,
    Parser,
    Range,
    Serial,
    XLError,
    arg,
    col_letters,
    fn,
    from_serial,
    num,
    nums,
    scalar,
    serial,
    text,
    truthy,
    values,
)


def numbers_only(a):
    out = nums(a)
    if not out:
        raise FormulaError(DIV0)
    return out


# ---------- finance ----------

def _annuity(rate, nper, pmt, pv, fv, kind):
    """The time-value-of-money equation; 0 when the arguments balance."""
    if rate == 0:
        return pv + pmt * nper + fv
    g = (1 + rate) ** nper
    return pv * g + pmt * (1 + rate * kind) * (g - 1) / rate + fv


@fn("PMT")
def _pmt(e, a):
    rate, nper, pv, fv, kind = num(a[0]), num(a[1]), num(a[2]), num(arg(a, 3, 0)), num(arg(a, 4, 0))
    if nper == 0:
        raise FormulaError(NUM)
    if rate == 0:
        return -(pv + fv) / nper
    g = (1 + rate) ** nper
    return -(rate * (fv + pv * g)) / ((1 + rate * kind) * (g - 1))


@fn("FV")
def _fv(e, a):
    rate, nper, pmt, pv, kind = num(a[0]), num(a[1]), num(a[2]), num(arg(a, 3, 0)), num(arg(a, 4, 0))
    return -_annuity(rate, nper, pmt, pv, 0, kind)


@fn("PV")
def _pv(e, a):
    rate, nper, pmt, fv, kind = num(a[0]), num(a[1]), num(a[2]), num(arg(a, 3, 0)), num(arg(a, 4, 0))
    if rate == 0:
        return -(fv + pmt * nper)
    return -(fv + pmt * (1 + rate * kind) * ((1 + rate) ** nper - 1) / rate) / (1 + rate) ** nper


@fn("NPER")
def _nper(e, a):
    rate, pmt, pv, fv, kind = num(a[0]), num(a[1]), num(a[2]), num(arg(a, 3, 0)), num(arg(a, 4, 0))
    if rate == 0:
        if pmt == 0:
            raise FormulaError(NUM)
        return -(pv + fv) / pmt
    top, bottom = pmt * (1 + rate * kind) - fv * rate, pmt * (1 + rate * kind) + pv * rate
    if top / bottom <= 0:
        raise FormulaError(NUM)
    return math.log(top / bottom) / math.log(1 + rate)


def _solve(f, guess):
    """Newton's method with a numeric derivative."""
    x = guess
    for _ in range(100):
        y = f(x)
        if abs(y) < 1e-10:
            return x
        h = 1e-7 * max(abs(x), 1)
        d = (f(x + h) - y) / h
        if d == 0:
            break
        x -= y / d
        if x <= -1:
            x = -0.999999
    raise FormulaError(NUM)


@fn("RATE")
def _rate(e, a):
    nper, pmt, pv, fv, kind = num(a[0]), num(a[1]), num(a[2]), num(arg(a, 3, 0)), num(arg(a, 4, 0))
    return _solve(lambda r: _annuity(r, nper, pmt, pv, fv, kind), num(arg(a, 5, 0.1)))


@fn("NPV")
def _npv(e, a):
    rate = num(a[0])
    return sum(v / (1 + rate) ** (i + 1) for i, v in enumerate(nums(a[1:])))


@fn("IRR")
def _irr(e, a):
    flows = nums([a[0]])
    if not (any(v > 0 for v in flows) and any(v < 0 for v in flows)):
        raise FormulaError(NUM)
    return _solve(lambda r: sum(v / (1 + r) ** i for i, v in enumerate(flows)), num(arg(a, 1, 0.1)))


# ---------- statistics ----------

@fn("STDEV", "STDEV.S")
def _stdev(e, a):
    v = numbers_only(a)
    if len(v) < 2:
        raise FormulaError(DIV0)
    return statistics.stdev(v)


@fn("STDEV.P", "STDEVP")
def _stdevp(e, a):
    return statistics.pstdev(numbers_only(a))


@fn("VAR", "VAR.S")
def _var(e, a):
    v = numbers_only(a)
    if len(v) < 2:
        raise FormulaError(DIV0)
    return statistics.variance(v)


@fn("VAR.P", "VARP")
def _varp(e, a):
    return statistics.pvariance(numbers_only(a))


@fn("MODE", "MODE.SNGL")
def _mode(e, a):
    v = nums(a)
    counts = Counter(v)
    best = max(counts.values(), default=0)
    if best < 2:
        raise FormulaError(NA)
    return float(next(x for x in v if counts[x] == best))


def _percentile(v, k):
    v = sorted(v)
    if not v or not 0 <= k <= 1:
        raise FormulaError(NUM)
    pos = k * (len(v) - 1)
    lo = math.floor(pos)
    hi = min(lo + 1, len(v) - 1)
    return v[lo] + (v[hi] - v[lo]) * (pos - lo)


@fn("PERCENTILE", "PERCENTILE.INC")
def _pct(e, a):
    return _percentile(nums([a[0]]), num(a[1]))


@fn("QUARTILE", "QUARTILE.INC")
def _quartile(e, a):
    q = int(num(a[1]))
    if not 0 <= q <= 4:
        raise FormulaError(NUM)
    return _percentile(nums([a[0]]), q / 4)


@fn("RANK", "RANK.EQ")
def _rank(e, a):
    x, v = num(a[0]), nums([a[1]])
    if x not in v:
        raise FormulaError(NA)
    ascending = truthy(arg(a, 2, 0))
    return float(sum(1 for y in v if (y < x if ascending else y > x)) + 1)


@fn("COUNTUNIQUE")
def _countunique(e, a):
    return float(len({text(v) for v in values(a) if v is not None and v != ""}))


# ---------- math ----------

@fn("SUMPRODUCT")
def _sumproduct(e, a):
    ranges = [x if isinstance(x, Range) else None for x in a]
    if not ranges or None in ranges:
        raise FormulaError(VALUE)
    h, w = ranges[0].h, ranges[0].w
    if any((r.h, r.w) != (h, w) for r in ranges):
        raise FormulaError(VALUE)
    total = 0.0
    for i in range(h):
        for j in range(w):
            p = 1.0
            for r in ranges:
                v = r.at(i, j)
                if isinstance(v, XLError):
                    raise FormulaError(v)
                p *= v if isinstance(v, (int, float)) and not isinstance(v, bool) else 0.0
            total += p
    return total


@fn("SUMSQ")
def _sumsq(e, a):
    return float(sum(v * v for v in nums(a)))


def _round_to(x, sig, up):
    if sig == 0:
        return 0.0
    q = x / sig
    return (math.ceil(q) if up else math.floor(q)) * sig


@fn("CEILING", "CEILING.MATH")
def _ceiling(e, a):
    return _round_to(num(a[0]), abs(num(arg(a, 1, 1))) or 1, up=True)


@fn("FLOOR", "FLOOR.MATH")
def _floor(e, a):
    return _round_to(num(a[0]), abs(num(arg(a, 1, 1))) or 1, up=False)


@fn("MROUND")
def _mround(e, a):
    x, m = num(a[0]), num(a[1])
    if m == 0:
        return 0.0
    if x * m < 0:
        raise FormulaError(NUM)
    return math.floor(x / m + 0.5) * m


@fn("EVEN")
def _even(e, a):
    x = num(a[0])
    n = math.ceil(abs(x) / 2) * 2
    return float(n if x >= 0 else -n)


@fn("ODD")
def _odd(e, a):
    x = num(a[0])
    n = math.ceil(abs(x))
    n = n if n % 2 else n + 1
    return float(n if x >= 0 else -n)


@fn("FACT")
def _fact(e, a):
    x = int(num(a[0]))
    if x < 0:
        raise FormulaError(NUM)
    return float(math.factorial(x))


@fn("GCD")
def _gcd(e, a):
    return float(math.gcd(*[int(v) for v in nums(a)]))


@fn("LCM")
def _lcm(e, a):
    return float(math.lcm(*[int(v) for v in nums(a)]))


for _name, _f in (("SIN", math.sin), ("COS", math.cos), ("TAN", math.tan), ("ATAN", math.atan),
                  ("DEGREES", math.degrees), ("RADIANS", math.radians)):
    fn(_name)(lambda e, a, f=_f: f(num(a[0])))


@fn("ASIN")
def _asin(e, a):
    x = num(a[0])
    if not -1 <= x <= 1:
        raise FormulaError(NUM)
    return math.asin(x)


@fn("ACOS")
def _acos(e, a):
    x = num(a[0])
    if not -1 <= x <= 1:
        raise FormulaError(NUM)
    return math.acos(x)


@fn("ATAN2")
def _atan2(e, a):
    x, y = num(a[0]), num(a[1])   # Excel: ATAN2(x, y)
    if x == 0 and y == 0:
        raise FormulaError(DIV0)
    return math.atan2(y, x)


# ---------- dates & times ----------

def _date(v):
    return from_serial(num(v))


def _add_months(d, months):
    y, m = divmod(d.month - 1 + months, 12)
    y, m = d.year + y, m + 1
    return d.replace(year=y, month=m, day=min(d.day, calendar.monthrange(y, m)[1]))


@fn("EDATE")
def _edate(e, a):
    return Serial(math.floor(serial(_add_months(_date(a[0]), int(num(a[1]))))))


@fn("EOMONTH")
def _eomonth(e, a):
    d = _add_months(_date(a[0]).replace(day=1), int(num(a[1])))
    return Serial(math.floor(serial(d.replace(day=calendar.monthrange(d.year, d.month)[1]))))


def _holidays(a, i):
    h = arg(a, i, None)
    if h is None:
        return set()
    return {math.floor(v) for v in nums([h])}


@fn("NETWORKDAYS")
def _networkdays(e, a):
    start, end = math.floor(num(a[0])), math.floor(num(a[1]))
    off = _holidays(a, 2)
    sign = 1 if end >= start else -1
    days = range(start, end + sign, sign)
    return float(sign * sum(1 for d in days if from_serial(d).weekday() < 5 and d not in off))


@fn("WORKDAY")
def _workday(e, a):
    d, left = math.floor(num(a[0])), int(num(a[1]))
    off = _holidays(a, 2)
    step = 1 if left >= 0 else -1
    while left:
        d += step
        if from_serial(d).weekday() < 5 and d not in off:
            left -= step
    return Serial(d)


@fn("DATEDIF")
def _datedif(e, a):
    s, t, unit = _date(a[0]), _date(a[1]), text(a[2]).upper()
    if t < s:
        raise FormulaError(NUM)
    months = (t.year - s.year) * 12 + t.month - s.month - (1 if t.day < s.day else 0)
    if unit == "Y":
        return float(months // 12)
    if unit == "M":
        return float(months)
    if unit == "D":
        return float((t - s).days)
    if unit == "YM":
        return float(months % 12)
    if unit == "MD":
        return float((t - _add_months(s, months)).days)
    if unit == "YD":
        start = s.replace(year=t.year)
        if start > t:
            start = s.replace(year=t.year - 1)
        return float((t - start).days)
    raise FormulaError(NUM)


@fn("HOUR")
def _hour(e, a):
    return float(_date(a[0]).hour)


@fn("MINUTE")
def _minute(e, a):
    return float(_date(a[0]).minute)


@fn("SECOND")
def _second(e, a):
    return float(_date(a[0]).second)


@fn("TIME")
def _time(e, a):
    secs = num(a[0]) * 3600 + num(a[1]) * 60 + num(a[2])
    return (secs % 86400) / 86400


@fn("DATEVALUE")
def _datevalue(e, a):
    """A date written as text -> its date serial ("15 Oct 2026", "10/15/2026", "2026-10-15"…; time dropped)."""
    v = num(text(a[0]))
    if v < 1:                                       # a time alone isn't a date
        raise FormulaError(VALUE)
    return Serial(math.floor(v))


@fn("TIMEVALUE")
def _timevalue(e, a):
    """A time written as text -> the fraction of a day ("2:30 PM", "14:30:15", "15 Oct 2026 6:00")."""
    v = num(text(a[0]))
    return v - math.floor(v)


# ---------- text ----------

@fn("IMAGE")
def _image(e, a):
    """IMAGE(source, [alt_text], [sizing], [height], [width]): a web picture shown in the cell (cell_images.py).
    sizing 0 fit (default), 1 fill, 2 original size, 3 custom height / width in pixels."""
    from cell_images import ImageValue

    def opt(i):
        return a[i] if len(a) > i and a[i] is not None and not (isinstance(a[i], str) and a[i] == "") else None
    url = text(a[0]).strip()
    if not url:
        raise FormulaError(VALUE)
    second = opt(1)
    if isinstance(second, (int, float)) and not isinstance(second, bool) and second in (1, 2, 3, 4) \
            and len(a) <= 4:
        # Google Sheets' IMAGE(url, mode, [height], [width]): 1 fit, 2 stretch, 3 original, 4 custom size
        height = num(opt(2)) if opt(2) is not None else None
        width = num(opt(3)) if opt(3) is not None else None
        if second == 4 and not (height and width and height > 0 and width > 0):
            raise FormulaError(VALUE)
        return ImageValue(url, "", int(second) - 1, height, width)
    alt = text(second) if second is not None else ""
    sizing = int(num(opt(2))) if opt(2) is not None else 0
    if sizing not in (0, 1, 2, 3):
        raise FormulaError(VALUE)
    height = num(opt(3)) if opt(3) is not None else None
    width = num(opt(4)) if opt(4) is not None else None
    if sizing == 3 and not (height and width and height > 0 and width > 0):
        raise FormulaError(VALUE)
    return ImageValue(url, alt, sizing, height, width)


@fn("CHAR")
def _char(e, a):
    n = int(num(a[0]))
    if not 1 <= n <= 255:
        raise FormulaError(VALUE)
    return bytes([n]).decode("cp1252", errors="replace")


@fn("CODE")
def _code(e, a):
    s = text(a[0])
    if not s:
        raise FormulaError(VALUE)
    try:
        return float(s[0].encode("cp1252")[0])
    except UnicodeEncodeError:
        return float(ord(s[0]))


@fn("CLEAN")
def _clean(e, a):
    return "".join(ch for ch in text(a[0]) if ord(ch) >= 32)


@fn("FIXED")
def _fixed(e, a):
    x, d, no_commas = num(a[0]), int(num(arg(a, 1, 2))), truthy(arg(a, 2, False))
    from formula import ROUND_HALF_UP, rnd
    v = rnd(x, d, ROUND_HALF_UP)
    s = f"{v:{'' if no_commas else ','}.{max(d, 0)}f}"
    return s


@fn("ISEVEN")
def _iseven(e, a):
    return int(abs(num(a[0]))) % 2 == 0


@fn("ISODD")
def _isodd(e, a):
    return int(abs(num(a[0]))) % 2 == 1


# ---------- references ----------

def _ref_of(e, node, sheet):
    """(sheet, r1, c1, r2, c2) of a reference argument (or of the cell being calculated)."""
    if node is None:
        if not e.stack:
            raise FormulaError(VALUE)
        s, r, c = e.stack[-1]
        return s, r, c, r, c
    if node[0] == "ref":
        _, name, r1, c1, r2, c2 = node
        return e.sheet_name(name, sheet), r1, c1, r2, c2
    if node[0] == "name":
        target = e.defined.get(node[1].lower())
        if target:
            return _ref_of(e, Parser(target.lstrip("=")).parse(), sheet)
    v = e.ev(node, sheet)
    if isinstance(v, Range):
        return v.sheet, v.r1, v.c1, v.r2, v.c2
    raise FormulaError(VALUE)


@fn("ROW", lazy="raw")
def _row(e, nodes, sheet):
    return float(_ref_of(e, nodes[0] if nodes else None, sheet)[1] + 1)


@fn("COLUMN", lazy="raw")
def _column(e, nodes, sheet):
    return float(_ref_of(e, nodes[0] if nodes else None, sheet)[2] + 1)


@fn("OFFSET", lazy="raw")
def _offset(e, nodes, sheet):
    if len(nodes) < 3:
        raise FormulaError(VALUE)
    s, r1, c1, r2, c2 = _ref_of(e, nodes[0], sheet)
    dr, dc = int(num(e.ev(nodes[1], sheet))), int(num(e.ev(nodes[2], sheet)))
    h = int(num(e.ev(nodes[3], sheet))) if len(nodes) > 3 and nodes[3][0] != "missing" else r2 - r1 + 1
    w = int(num(e.ev(nodes[4], sheet))) if len(nodes) > 4 and nodes[4][0] != "missing" else c2 - c1 + 1
    r, c = r1 + dr, c1 + dc
    if r < 0 or c < 0 or h < 1 or w < 1:
        raise FormulaError(REF)
    if h == 1 and w == 1:
        return e.cell(s, r, c)
    return Range(e, s, r, c, r + h - 1, c + w - 1)


@fn("INDIRECT")
def _indirect(e, a):
    ref = text(scalar(a[0])).strip()
    try:
        node = Parser(ref.lstrip("=")).parse()
    except FormulaError:
        raise FormulaError(REF) from None
    if node[0] not in ("ref", "name"):
        raise FormulaError(REF)
    sheet = e.stack[-1][0] if e.stack else next(iter(e.book))
    return e.ev(node, sheet)


@fn("ADDRESS")
def _address(e, a):
    r, c, kind = int(num(a[0])), int(num(a[1])), int(num(arg(a, 2, 1)))
    if r < 1 or c < 1:
        raise FormulaError(VALUE)
    col, row = col_letters(c - 1), str(r)
    ref = {1: f"${col}${row}", 2: f"{col}${row}", 3: f"${col}{row}", 4: f"{col}{row}"}.get(kind)
    if ref is None:
        raise FormulaError(VALUE)
    sheet = arg(a, 4, None)
    if sheet is not None:
        from formula import quote_sheet
        ref = f"{quote_sheet(text(sheet))}!{ref}"
    return ref
