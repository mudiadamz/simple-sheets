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
