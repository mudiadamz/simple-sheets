"""Insert / delete whole rows or columns in the model, fixing styles, sizes, merges and formulas."""
from formula import adjust


def _move(i, at, n):
    """New index after inserting (n > 0) or deleting (n < 0) at `at`; None if deleted."""
    if i < at:
        return i
    if n > 0:
        return i + n
    return None if i < at - n else i + n


def _frozen(count, at, n):
    """Frozen row/column count after inserting (n > 0) or deleting (n < 0) at `at`."""
    if n > 0:
        return count + n if at < count else count
    return count - len([i for i in range(at, at - n) if i < count])


def _span(i1, i2, at, n):
    a, b = _move(i1, at, n), _move(i2, at, n)
    if a is None and b is None:
        return None
    a = at if a is None else a
    b = at - 1 if b is None else b
    return (a, b) if a <= b else None


def _rekey(d, axis, at, n):
    """Shift dict keys (r, c) or plain indexes; drop deleted ones."""
    out = {}
    for k, v in d.items():
        if isinstance(k, tuple):
            r, c = k
            i = _move(r if axis == "row" else c, at, n)
            if i is not None:
                out[(i, c) if axis == "row" else (r, i)] = v
        else:
            i = _move(k, at, n)
            if i is not None:
                out[i] = v
    return out


def _remap(m, tail, at, n):
    """Update a row/column map (model index -> index in the file on disk)."""
    m = list(m)
    while len(m) < at - min(n, 0):   # make the map cover the touched area
        m.append(tail)
        tail += 1
    if n > 0:
        m[at:at] = [None] * n
    else:
        del m[at:at - n]
    return m, tail


def change(book, name, axis, at, n):
    """Insert (n > 0) or delete (n < 0) rows/columns at `at` on sheet `name`."""
    s = book[name]
    if axis == "row":
        if n > 0:
            if at <= len(s.rows):
                s.rows[at:at] = [[] for _ in range(n)]
        else:
            del s.rows[at:at - n]
        s.heights = _rekey(s.heights, None, at, n)
        s.outline_rows = _rekey(s.outline_rows, None, at, n)
        s.hidden_rows = {j for j in (_move(i, at, n) for i in s.hidden_rows) if j is not None}
        s.freeze = (_frozen(s.freeze[0], at, n), s.freeze[1])
        s.filtered_rows = {j for j in (_move(i, at, n) for i in s.filtered_rows) if j is not None}
        s.rowmap, s.rowtail = _remap(s.rowmap, s.rowtail, at, n)
    else:
        for row in s.rows:
            if n > 0:
                if at <= len(row):
                    row[at:at] = [""] * n
            else:
                del row[at:at - n]
        s.widths = _rekey(s.widths, None, at, n)
        s.outline_cols = _rekey(s.outline_cols, None, at, n)
        s.px = _rekey(s.px, None, at, n)
        s.hidden_cols = {j for j in (_move(i, at, n) for i in s.hidden_cols) if j is not None}
        s.freeze = (s.freeze[0], _frozen(s.freeze[1], at, n))
        s.colmap, s.coltail = _remap(s.colmap, s.coltail, at, n)
    s.styles = _rekey(s.styles, axis, at, n)
    s.cache = _rekey(s.cache, axis, at, n)
    s.links = _rekey(s.links, axis, at, n)
    s.rich = _rekey(s.rich, axis, at, n)
    s.comments = _rekey(s.comments, axis, at, n)
    s.threads = _rekey(s.threads, axis, at, n)
    if s.filter:   # the filter range (and its columns' criteria) follow
        r1, c1, r2, c2 = s.filter["range"]
        sp = _span(r1, r2, at, n) if axis == "row" else _span(c1, c2, at, n)
        if sp is None or (axis == "row" and _move(r1, at, n) is None):
            s.filter, s.filtered_rows = None, set()          # range (or its header row) deleted
        else:
            rng = (sp[0], c1, sp[1], c2) if axis == "row" else (r1, sp[0], r2, sp[1])
            crit = s.filter["criteria"] if axis == "row" else _rekey(s.filter["criteria"], None, at, n)
            s.filter = {"range": rng, "criteria": crit}
    merges = []
    for r1, c1, r2, c2 in s.merges:
        sp = _span(r1, r2, at, n) if axis == "row" else _span(c1, c2, at, n)
        if sp:
            m = (sp[0], c1, sp[1], c2) if axis == "row" else (r1, sp[0], r2, sp[1])
            if m[0] < m[2] or m[1] < m[3]:
                merges.append(m)
    s.merges = merges
    s.view = None
    # formulas anywhere that point at this sheet
    changed = []
    for home, sh in book.items():
        for r, row in enumerate(sh.rows):
            for c, t in enumerate(row):
                if t.startswith("=") and len(t) > 1:
                    new = adjust(t, home, name, axis, at, n)
                    if new != t:
                        row[c] = new
                        changed.append((home, r, c))
    return changed


# ---------- move whole rows / columns (drag a header) ----------

def mover(a, n, t):
    """Lines a .. a+n-1 go before line t (t outside the block): -> perm(i) = new index of line i."""
    if t < a:
        def perm(i):
            if a <= i < a + n:
                return t + i - a
            return i + n if t <= i < a else i
    else:
        def perm(i):
            if a <= i < a + n:
                return t - n + i - a
            return i - n if a + n <= i < t else i
    perm.span = (min(a, t), max(a + n, t) - 1)   # every line that changes place: ranges covering it all stay
    return perm


def _perm_keys(d, axis, perm):
    out = {}
    for k, v in d.items():
        if isinstance(k, tuple):
            r, c = k
            out[(perm(r), c) if axis == "row" else (r, perm(c))] = v
        else:
            out[perm(k)] = v
    return out


def _perm_list(items, perm, size, fill):
    """Reorder a list (padded to `size` with fill()); trailing fills trimmed again."""
    items = list(items) + [fill() for _ in range(max(size - len(items), 0))]
    out = [None] * len(items)
    for i, v in enumerate(items):
        out[perm(i)] = v
    while out and out[-1] in ([], ""):
        out.pop()
    return out


def move(book, name, axis, a, n, t):
    """Move rows / columns a .. a+n-1 of sheet `name` so they sit before line t (index before the move).
    Values, formats, sizes, hidden state, comments, links and merged cells move with them; formulas on
    every sheet follow the moved cells. Returns the changed formula cells. Raises ValueError when a merged
    block would be split."""
    from formula import permute
    if a <= t <= a + n:
        return []
    s = book[name]
    perm = mover(a, n, t)
    lo, hi = min(a, t), max(a + n, t)
    for r1, c1, r2, c2 in s.merges:                       # a merged block must move as a whole
        i1, i2 = (r1, r2) if axis == "row" else (c1, c2)
        new = sorted(perm(i) for i in range(i1, i2 + 1))
        if new != list(range(new[0], new[0] + len(new))):
            raise ValueError("merged cells would be split")
    size = hi
    if axis == "row":
        s.rows = _perm_list(s.rows, perm, size, list)
        s.heights = _perm_keys(s.heights, None, perm)
        s.outline_rows = _perm_keys(s.outline_rows, None, perm)
        s.hidden_rows = {perm(i) for i in s.hidden_rows}
        s.filtered_rows = {perm(i) for i in s.filtered_rows}
        m = list(s.rowmap)
        while len(m) < size:
            m.append(s.rowtail)
            s.rowtail += 1
        s.rowmap = [m[i] for i in sorted(range(len(m)), key=perm)]   # model row -> row on disk, reordered
    else:
        for k, row in enumerate(s.rows):
            if len(row) > lo:
                s.rows[k] = _perm_list(row, perm, size, str)
        s.widths = _perm_keys(s.widths, None, perm)
        s.outline_cols = _perm_keys(s.outline_cols, None, perm)
        s.px = _perm_keys(s.px, None, perm)
        s.hidden_cols = {perm(i) for i in s.hidden_cols}
        m = list(s.colmap)
        while len(m) < size:
            m.append(s.coltail)
            s.coltail += 1
        s.colmap = [m[i] for i in sorted(range(len(m)), key=perm)]
        if s.filter:
            s.filter = {**s.filter, "criteria": _perm_keys(s.filter["criteria"], None, perm)}
    for f in ("styles", "cache", "links", "rich", "comments", "threads"):
        setattr(s, f, _perm_keys(getattr(s, f), axis, perm))
    merges = []
    for r1, c1, r2, c2 in s.merges:
        if axis == "row":
            p1, p2 = sorted((perm(r1), perm(r2)))
            merges.append((p1, c1, p2, c2))
        else:
            p1, p2 = sorted((perm(c1), perm(c2)))
            merges.append((r1, p1, r2, p2))
    s.merges = merges
    s.view = None
    changed = []
    for home, sh in book.items():
        for r, row in enumerate(sh.rows):
            for c, t_ in enumerate(row):
                if t_.startswith("=") and len(t_) > 1:
                    new = permute(t_, home, name, axis, perm)
                    if new != t_:
                        row[c] = new
                        changed.append((home, r, c))
    return changed


FIELDS = ("rows", "styles", "heights", "widths", "px", "merges", "cache",
          "rowmap", "rowtail", "colmap", "coltail", "hidden_rows", "hidden_cols", "hidden", "links", "freeze", "filter",
          "filtered_rows", "comments", "outline_rows", "outline_cols", "rich", "threads")


def _copy(f, v):
    if f == "rows":
        return [list(r) for r in v]
    return type(v)(v) if isinstance(v, (dict, list, set)) else v


def snapshot(book):
    """Copy of everything a structural change can touch (rows copied one level deep)."""
    return {name: {f: _copy(f, getattr(s, f)) for f in FIELDS} for name, s in book.items()}


def restore(book, snap):
    for name, st in snap.items():
        s = book[name]
        for f in FIELDS:
            setattr(s, f, _copy(f, st[f]))
        s.view = None


def _compose(m, tail, cur_m, cur_tail):
    """Map of a past state -> old file, re-expressed against the new file (= current model)."""
    orig = lambda i: m[i] if i < len(m) else tail + i - len(m)
    inv = {v: i for i, v in enumerate(cur_m) if v is not None}
    new = lambda o: None if o is None else (inv.get(o) if o < cur_tail else o - cur_tail + len(cur_m))
    n = max(len(m), len(cur_m), cur_tail - tail + len(m), 0)
    return [new(orig(i)) for i in range(n)], new(orig(n))


def rebase(snap, book):
    """Call on every stored snapshot right before a save resets the maps."""
    for name, st in snap.items():
        s = book.get(name)
        if s is None:
            continue
        st["rowmap"], st["rowtail"] = _compose(st["rowmap"], st["rowtail"], s.rowmap, s.rowtail)
        st["colmap"], st["coltail"] = _compose(st["colmap"], st["coltail"], s.colmap, s.coltail)


# ---------- insert / delete cells (shift part of the grid) ----------

def shift_cells(sheet, rect, direction):
    """-> [(kind, r, c, new)] changes that shift cells of the block, like Excel's
    Insert cells (direction 'down' / 'right') or Delete cells ('up' / 'left')."""
    r1, c1, r2, c2 = rect
    h, w = r2 - r1 + 1, c2 - c1 + 1
    get = lambda r, c: sheet.rows[r][c] if r < len(sheet.rows) and c < len(sheet.rows[r]) else ""
    sty = lambda r, c: sheet.styles.get((r, c))
    out = []
    if direction in ("down", "up"):
        last = max(len(sheet.rows) - 1, r2) + h
        for c in range(c1, c2 + 1):
            for r in range(r1, last + 1):
                src = r - h if direction == "down" else r + h
                ok = src >= r1 if direction == "down" else True
                out += [("v", r, c, get(src, c) if ok else ""), ("s", r, c, sty(src, c) if ok else None)]
    else:
        for r in range(r1, r2 + 1):
            last = max(len(sheet.rows[r]) if r < len(sheet.rows) else 0, c2 + 1) + w
            for c in range(c1, last):
                src = c - w if direction == "right" else c + w
                ok = src >= c1 if direction == "right" else True
                out += [("v", r, c, get(r, src) if ok else ""), ("s", r, c, sty(r, src) if ok else None)]
    return out
