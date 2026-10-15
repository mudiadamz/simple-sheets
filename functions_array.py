"""Dynamic-array functions (results spill into the cells below / right): SEQUENCE, FILTER, UNIQUE, SORT,
SORTBY, TRANSPOSE, RANDARRAY, VSTACK, HSTACK, TAKE, DROP, CHOOSEROWS, CHOOSECOLS, TOCOL, TOROW, TEXTSPLIT,
WRAPROWS, WRAPCOLS, MMULT, FREQUENCY. Registered into formula.FUNCS."""
import functools
import random

from formula import (
    CALC,
    NA,
    VALUE,
    Array,
    FormulaError,
    Range,
    XLError,
    arg,
    as_rows,
    compare,
    fn,
    num,
    nums,
    text,
    truthy,
)


def rows_of(v):
    return as_rows(v)


def result(rows):
    if not rows or not rows[0]:
        raise FormulaError(CALC)          # Excel: an empty array
    return Array(rows)


@fn("SEQUENCE")
def _sequence(e, a):
    h, w = int(num(a[0])), int(num(arg(a, 1, 1)))
    start, step = num(arg(a, 2, 1)), num(arg(a, 3, 1))
    if h < 1 or w < 1 or h * w > 1_000_000:
        raise FormulaError(VALUE)
    return Array([[start + (i * w + j) * step for j in range(w)] for i in range(h)])


@fn("RANDARRAY")
def _randarray(e, a):
    h, w = int(num(arg(a, 0, 1))), int(num(arg(a, 1, 1)))
    lo, hi, whole = num(arg(a, 2, 0)), num(arg(a, 3, 1)), truthy(arg(a, 4, False))
    if h < 1 or w < 1 or lo > hi:
        raise FormulaError(VALUE)
    pick = (lambda: float(random.randint(int(lo), int(hi)))) if whole else (lambda: random.uniform(lo, hi))
    return Array([[pick() for _ in range(w)] for _ in range(h)])


@fn("TRANSPOSE")
def _transpose(e, a):
    rows = rows_of(a[0])
    return Array([list(col) for col in zip(*rows)])


@fn("FILTER")
def _filter(e, a):
    rows = rows_of(a[0])
    keep = rows_of(a[1])
    if len(keep) == len(rows) and len(keep[0]) == 1:          # one flag per row
        out = [r for r, k in zip(rows, keep) if truthy(k[0])]
    elif len(keep) == 1 and len(keep[0]) == len(rows[0]):     # one flag per column
        idx = [j for j, k in enumerate(keep[0]) if truthy(k)]
        out = [[r[j] for j in idx] for r in rows] if idx else []
    else:
        raise FormulaError(VALUE)
    if not out or not out[0]:
        if len(a) > 2:
            return a[2]
        raise FormulaError(CALC)
    return Array(out)


def _key(v):
    return v.casefold() if isinstance(v, str) else v


@fn("UNIQUE")
def _unique(e, a):
    rows = rows_of(a[0])
    by_col, once = truthy(arg(a, 1, False)), truthy(arg(a, 2, False))
    items = [list(c) for c in zip(*rows)] if by_col else rows
    counts = {}
    for it in items:
        k = tuple(_key(v) for v in it)
        counts[k] = counts.get(k, 0) + 1
    seen, out = set(), []
    for it in items:
        k = tuple(_key(v) for v in it)
        if k in seen or (once and counts[k] > 1):
            continue
        seen.add(k)
        out.append(it)
    if by_col:
        out = [list(r) for r in zip(*out)] if out else []
    return result(out)


def _sort_rows(rows, keys):
    """keys: [(column index, ascending)]"""
    def cmp(x, y):
        for j, asc in keys:
            try:
                c = compare(x[j], y[j])
            except FormulaError:
                c = 0
            if c:
                return c if asc else -c
        return 0
    return sorted(rows, key=functools.cmp_to_key(cmp))


@fn("SORT")
def _sort(e, a):
    rows = rows_of(a[0])
    idx, order, by_col = int(num(arg(a, 1, 1))), num(arg(a, 2, 1)), truthy(arg(a, 3, False))
    items = [list(c) for c in zip(*rows)] if by_col else rows
    if not 1 <= idx <= len(items[0]):
        raise FormulaError(VALUE)
    out = _sort_rows(items, [(idx - 1, order >= 0)])
    return Array([list(r) for r in zip(*out)] if by_col else out)


@fn("SORTBY")
def _sortby(e, a):
    rows = rows_of(a[0])
    keyed = [list(r) + [] for r in rows]
    keys = []
    for k in range(1, len(a), 2):
        col = rows_of(a[k])
        if len(col) != len(rows):
            raise FormulaError(VALUE)
        for r, v in zip(keyed, col):
            r.append(v[0])
        keys.append((len(keyed[0]) - 1, num(arg(a, k + 1, 1)) >= 0))
    out = _sort_rows(keyed, keys)
    return Array([r[:len(rows[0])] for r in out])


@fn("VSTACK")
def _vstack(e, a):
    blocks = [rows_of(x) for x in a]
    w = max(len(b[0]) for b in blocks)
    return Array([r + [NA] * (w - len(r)) for b in blocks for r in b])


@fn("HSTACK")
def _hstack(e, a):
    blocks = [rows_of(x) for x in a]
    h = max(len(b) for b in blocks)
    out = []
    for i in range(h):
        row = []
        for b in blocks:
            row += b[i] if i < len(b) else [NA] * len(b[0])
        out.append(row)
    return Array(out)


@fn("TAKE")
def _take(e, a):
    rows = rows_of(a[0])
    n_r, n_c = arg(a, 1, None), arg(a, 2, None)
    if n_r is not None:
        n = int(num(n_r))
        rows = rows[:n] if n >= 0 else rows[n:]
    if n_c is not None:
        n = int(num(n_c))
        rows = [r[:n] if n >= 0 else r[n:] for r in rows]
    return result(rows)


@fn("DROP")
def _drop(e, a):
    rows = rows_of(a[0])
    n_r, n_c = arg(a, 1, None), arg(a, 2, None)
    if n_r is not None:
        n = int(num(n_r))
        rows = rows[n:] if n >= 0 else rows[:n]
    if n_c is not None:
        n = int(num(n_c))
        rows = [r[n:] if n >= 0 else r[:n] for r in rows]
    return result(rows)


def _pick(items, idx):
    out = []
    for i in idx:
        i = int(i)
        if i == 0 or abs(i) > len(items):
            raise FormulaError(VALUE)
        out.append(items[i - 1] if i > 0 else items[i])
    return out


@fn("CHOOSEROWS")
def _chooserows(e, a):
    return Array(_pick(rows_of(a[0]), nums(a[1:])))


@fn("CHOOSECOLS")
def _choosecols(e, a):
    cols = [list(c) for c in zip(*rows_of(a[0]))]
    return Array([list(r) for r in zip(*_pick(cols, nums(a[1:])))])


def _flat(rows, by_col, ignore):
    items = [v for c in zip(*rows) for v in c] if by_col else [v for r in rows for v in r]
    if ignore in (1, 3):
        items = [v for v in items if v is not None and v != ""]
    if ignore in (2, 3):
        items = [v for v in items if not isinstance(v, XLError)]
    return items


@fn("TOCOL")
def _tocol(e, a):
    items = _flat(rows_of(a[0]), truthy(arg(a, 2, False)), int(num(arg(a, 1, 0))))
    return result([[v] for v in items])


@fn("TOROW")
def _torow(e, a):
    items = _flat(rows_of(a[0]), truthy(arg(a, 2, False)), int(num(arg(a, 1, 0))))
    return result([items])


def _wrap(items, n, pad):
    if n < 1:
        raise FormulaError(VALUE)
    out = [items[i:i + n] for i in range(0, len(items), n)]
    if out and len(out[-1]) < n:
        out[-1] += [pad] * (n - len(out[-1]))
    return out


@fn("WRAPROWS")
def _wraprows(e, a):
    items = _flat(rows_of(a[0]), False, 0)
    return result(_wrap(items, int(num(a[1])), arg(a, 2, NA)))


@fn("WRAPCOLS")
def _wrapcols(e, a):
    items = _flat(rows_of(a[0]), False, 0)
    return result([list(r) for r in zip(*_wrap(items, int(num(a[1])), arg(a, 2, NA)))])


@fn("TEXTSPLIT")
def _textsplit(e, a):
    s = text(a[0])
    col_sep, row_sep = arg(a, 1, None), arg(a, 2, None)
    ignore = truthy(arg(a, 3, False))
    seps = lambda v: [text(x) for x in (sum(rows_of(v), []) if isinstance(v, Range) else [v])] if v is not None else []  # noqa: E731

    def split(t, ss):
        parts = [t]
        for sep in ss:
            if sep:
                parts = [p for q in parts for p in q.split(sep)]
        return [p for p in parts if p] if ignore else parts
    lines = split(s, seps(row_sep)) if row_sep is not None else [s]
    rows = [split(line, seps(col_sep)) for line in lines]
    w = max(len(r) for r in rows)
    return result([r + [NA] * (w - len(r)) for r in rows])


@fn("MMULT")
def _mmult(e, a):
    x, y = rows_of(a[0]), rows_of(a[1])
    if len(x[0]) != len(y):
        raise FormulaError(VALUE)
    return Array([[sum(num(x[i][k]) * num(y[k][j]) for k in range(len(y))) for j in range(len(y[0]))]
                  for i in range(len(x))])


@fn("FREQUENCY")
def _frequency(e, a):
    data, bins = sorted(nums([a[0]])), sorted(nums([a[1]]))
    out, i = [], 0
    for b in bins:
        n = 0
        while i < len(data) and data[i] <= b:
            n, i = n + 1, i + 1
        out.append([float(n)])
    out.append([float(len(data) - i)])
    return Array(out)
