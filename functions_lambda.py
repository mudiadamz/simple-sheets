"""LET, LAMBDA and the functions that take one (MAP, BYROW, BYCOL, REDUCE, SCAN, MAKEARRAY), plus XMATCH,
SWITCH, TEXTBEFORE / TEXTAFTER, WEEKNUM / ISOWEEKNUM. Registered into formula.FUNCS."""
import math

from formula import NA, VALUE, Array, FormulaError, Lambda, Range, XLError, arg, as_rows, compare, fn, from_serial, num, text


def _name_of(node):
    if node[0] == "name":
        return node[1].lower()
    if node[0] == "ref" and node[1] is None:          # a short name like "x1" lexes as a cell reference
        raise FormulaError(VALUE)
    raise FormulaError(VALUE)


@fn("LET", lazy="raw")
def _let(e, nodes, sheet):
    if len(nodes) < 3 or len(nodes) % 2 == 0:
        raise FormulaError(VALUE)
    scope = {}
    e.env.append(scope)
    try:
        for i in range(0, len(nodes) - 1, 2):
            scope[_name_of(nodes[i])] = e.ev(nodes[i + 1], sheet)   # later values can use earlier names
        return e.ev(nodes[-1], sheet)
    finally:
        e.env.pop()


@fn("LAMBDA", lazy="raw")
def _lambda(e, nodes, sheet):
    if not nodes:
        raise FormulaError(VALUE)
    params = [_name_of(n) for n in nodes[:-1]]
    scope = {}
    for s in e.env:
        scope.update(s)
    return Lambda(params, nodes[-1], scope)


def _fn_arg(v):
    if not isinstance(v, Lambda):
        raise FormulaError(VALUE)
    return v


def _plain(v):
    if isinstance(v, Range):
        return v.at(0, 0) if v.h == 1 and v.w == 1 else v
    return v


@fn("MAP")
def _map(e, a):
    f = _fn_arg(a[-1])
    blocks = [as_rows(x) for x in a[:-1]]
    rows = blocks[0]
    return Array([[_plain(e.call_lambda(f, None, None, [b[i][j] for b in blocks])) for j in range(len(rows[0]))]
                  for i in range(len(rows))])


@fn("BYROW")
def _byrow(e, a):
    f = _fn_arg(a[1])
    return Array([[_plain(e.call_lambda(f, None, None, [Array([row])]))] for row in as_rows(a[0])])


@fn("BYCOL")
def _bycol(e, a):
    f = _fn_arg(a[1])
    cols = [list(c) for c in zip(*as_rows(a[0]))]
    return Array([[_plain(e.call_lambda(f, None, None, [Array([[v] for v in col])])) for col in cols]])


@fn("REDUCE")
def _reduce(e, a):
    acc, f = a[0], _fn_arg(a[2])
    for row in as_rows(a[1]):
        for v in row:
            acc = _plain(e.call_lambda(f, None, None, [acc, v]))
    return acc


@fn("SCAN")
def _scan(e, a):
    acc, f = a[0], _fn_arg(a[2])
    out = []
    for row in as_rows(a[1]):
        line = []
        for v in row:
            acc = _plain(e.call_lambda(f, None, None, [acc, v]))
            line.append(acc)
        out.append(line)
    return Array(out)


@fn("MAKEARRAY")
def _makearray(e, a):
    h, w, f = int(num(a[0])), int(num(a[1])), _fn_arg(a[2])
    if h < 1 or w < 1:
        raise FormulaError(VALUE)
    return Array([[_plain(e.call_lambda(f, None, None, [float(i + 1), float(j + 1)])) for j in range(w)]
                  for i in range(h)])


@fn("XMATCH")
def _xmatch(e, a):
    want, rows = a[0], as_rows(a[1])
    items = [r[0] for r in rows] if len(rows[0]) == 1 else rows[0]
    mode, search = int(num(arg(a, 2, 0))), int(num(arg(a, 3, 1)))
    order = range(len(items)) if search >= 0 else range(len(items) - 1, -1, -1)
    best = None
    for i in order:
        v = items[i]
        if v is None or isinstance(v, XLError):
            continue
        try:
            c = compare(v, want)
        except FormulaError:
            continue
        if c == 0:
            return float(i + 1)
        if mode == -1 and c < 0 and (best is None or compare(v, items[best]) > 0):
            best = i
        if mode == 1 and c > 0 and (best is None or compare(v, items[best]) < 0):
            best = i
        if mode == 2 and isinstance(want, str) and isinstance(v, str):
            import fnmatch
            if fnmatch.fnmatchcase(v.lower(), want.lower().replace("~*", "[*]").replace("~?", "[?]")):
                return float(i + 1)
    if best is None:
        raise FormulaError(NA)
    return float(best + 1)


@fn("SWITCH", lazy=True)
def _switch(e, a):
    value = a[0]()
    rest = a[1:]
    for i in range(0, len(rest) - 1, 2):
        try:
            if compare(value, rest[i]()) == 0:
                return rest[i + 1]()
        except FormulaError:
            continue
    if len(rest) % 2:
        return rest[-1]()
    raise FormulaError(NA)


def _split_at(s, delim, n, after):
    if not delim:
        return s if after else ""
    parts, idx, start = [], -1, 0
    while True:
        idx = s.find(delim, start)
        if idx < 0:
            break
        parts.append(idx)
        start = idx + len(delim)
    if not parts:
        raise FormulaError(NA)
    k = n - 1 if n > 0 else len(parts) + n
    if not 0 <= k < len(parts):
        raise FormulaError(NA)
    i = parts[k]
    return s[i + len(delim):] if after else s[:i]


@fn("TEXTBEFORE")
def _textbefore(e, a):
    return _split_at(text(a[0]), text(a[1]), int(num(arg(a, 2, 1))), after=False)


@fn("TEXTAFTER")
def _textafter(e, a):
    return _split_at(text(a[0]), text(a[1]), int(num(arg(a, 2, 1))), after=True)


@fn("WEEKNUM")
def _weeknum(e, a):
    d = from_serial(num(a[0]))
    start = int(num(arg(a, 1, 1)))
    first_day = 6 if start == 1 else 0                     # 1: weeks start Sunday, 2: Monday
    jan1 = d.replace(month=1, day=1)
    offset = (jan1.weekday() - first_day) % 7
    return float(math.floor((d.timetuple().tm_yday - 1 + offset) / 7) + 1)


@fn("ISOWEEKNUM")
def _isoweeknum(e, a):
    return float(from_serial(num(a[0])).isocalendar()[1])


@fn("ISOMITTED", lazy="raw")
def _isomitted(e, nodes, sheet):
    return nodes[0][0] == "missing" if nodes else True
