"""Data menu tools: text to columns, remove duplicates, flash fill (Ctrl+E), pivot table, goal seek."""
import re
import tkinter as tk
from tkinter import messagebox, simpledialog, ttk

from formatting import parse_number
from formula import col_letters


def _region(app):
    """Selection, or the block around the active cell when only one cell is selected."""
    g = app.table
    if g.multi():
        return g.sel_data()
    return g.region(*g.cur)


# ---------- text to columns ----------

DELIMS = {"Comma": ",", "Semicolon": ";", "Tab": "\t", "Space": " ", "Pipe": "|"}


def split_column(app, rect, delim, merge_runs=False):
    """Split the text of column rect[1] (rows rect[0]..rect[2]) into the columns to its right."""
    g = app.table
    r1, c, r2, _ = rect
    changes = []
    for r in range(r1, r2 + 1):
        t = g.text(r, c)
        if not t or t.startswith("="):
            continue
        parts = re.split("(?:" + re.escape(delim) + ")+", t) if merge_runs else t.split(delim)
        for k, p in enumerate(parts):
            changes.append((r, c + k, p.strip() if delim != "\t" else p))
    if not changes:
        return 0
    if any(g.text(r, cc) for r, cc, _ in changes if cc != c) and not messagebox.askokcancel(
            "Text to columns", "Cells to the right will be replaced. Continue?", parent=app):
        return 0
    app.set_cells(changes)
    return len(changes)


def text_to_columns(app):
    g = app.table
    r1, c1, r2, _ = g.sel_data() if g.multi() else g.region(*g.cur)
    choice = simpledialog.askstring("Text to columns", "Split at (comma, semicolon, tab, space, pipe, "
                                    "or type the character):", initialvalue="comma", parent=app)
    if not choice:
        return 0
    delim = DELIMS.get(choice.strip().capitalize(), choice)
    return split_column(app, (r1, g.cur[1], r2, g.cur[1]), delim, merge_runs=delim == " ")


# ---------- remove duplicates ----------

def dedupe(app, rect, header=True, cols=None):
    """Remove rows whose values (in `cols`, default all) repeat an earlier row; rows below move up
    inside the block. Returns how many rows were removed."""
    g = app.table
    r1, c1, r2, c2 = rect
    start = r1 + 1 if header else r1
    cols = cols if cols is not None else list(range(c1, c2 + 1))
    seen, keep = set(), []
    for r in range(start, r2 + 1):
        key = tuple(g.text(r, c).strip().casefold() for c in cols)
        if key in seen:
            continue
        seen.add(key)
        keep.append([g.text(r, c) for c in range(c1, c2 + 1)])
    removed = (r2 - start + 1) - len(keep)
    if removed:
        changes = []
        for i, r in enumerate(range(start, r2 + 1)):
            row = keep[i] if i < len(keep) else [""] * (c2 - c1 + 1)
            changes += [(r, c1 + j, v) for j, v in enumerate(row)]
        app.set_cells(changes)
    return removed


def remove_duplicates(app):
    rect = _region(app)
    first = [app.table.text(rect[0], c) for c in range(rect[1], rect[3] + 1)]
    header = bool(first) and all(t and parse_number(t) is None for t in first)
    n = dedupe(app, rect, header)
    messagebox.showinfo("Remove duplicates", f"{n} duplicate row{'s' if n != 1 else ''} removed." if n
                        else "No duplicates found.", parent=app)
    return n


# ---------- flash fill ----------

TOKEN = re.compile(r"[A-Za-z0-9À-￿]+")
CASES = {"": lambda s: s, "upper": str.upper, "lower": str.lower, "title": str.title}


def _token(s, k, f):
    t = TOKEN.findall(s)
    return f(t[k]) if -len(t) <= k < len(t) else None


def _pair(s, i, j, sep, f):
    t = TOKEN.findall(s)
    return f(t[i] + sep + t[j]) if len(t) > max(i, j) else None


def _programs(src):
    """Candidate ways to make an output from `src` (name, function)."""
    toks = TOKEN.findall(src)
    out = []
    for case, f in CASES.items():
        for k in range(-3, 4):
            out.append((("tok", k, case), lambda s, k=k, f=f: _token(s, k, f)))
        out.append((("initials", case), lambda s, f=f: f("".join(t[0] for t in TOKEN.findall(s)))))
        out.append((("whole", case), lambda s, f=f: f(s)))
    for i in range(min(len(toks), 4)):
        for j in range(min(len(toks), 4)):
            if i != j:
                for sep in (" ", ", ", ".", "_", "-", ""):
                    for case, f in CASES.items():
                        out.append((("pair", i, j, sep, case),
                                    lambda s, i=i, j=j, sep=sep, f=f: _pair(s, i, j, sep, f)))
    for d in ("@", ",", "-", "/", " ", ".", ":"):
        out.append((("before", d), lambda s, d=d: s.split(d, 1)[0] if d in s else None))
        out.append((("after", d), lambda s, d=d: s.split(d, 1)[1] if d in s else None))
    return out


def learn(examples):
    """[(source text, wanted)] -> function source -> output, or None when no simple rule fits."""
    if not examples:
        return None
    for _name, f in _programs(examples[0][0]):
        try:
            if all(f(s) == want for s, want in examples):
                return f
        except (IndexError, AttributeError):
            continue
    return None


def flash_fill(app):
    """Fill the active column below the typed examples, following the pattern in them (Ctrl+E)."""
    g = app.table
    r0, c = g.cur
    rows = len(g.sheet.rows)
    top = r0
    while top > 0 and g.text(top - 1, c):
        top -= 1
    examples_rows = [r for r in range(top, r0 + 1) if g.text(r, c)]
    for src_c in [c - 1, c + 1] + [x for x in range(c - 2, -1, -1)]:
        if src_c < 0:
            continue
        ex = [(g.text(r, src_c), g.text(r, c)) for r in examples_rows if g.text(r, src_c)]
        if len(ex) != len(examples_rows) or not ex:
            continue
        f = learn(ex) or (learn(ex[1:]) if len(ex) > 1 else None)   # the first row may be a header
        if f is None:
            continue
        changes, r = [], (examples_rows[-1] + 1 if examples_rows else r0)
        while r < rows and g.text(r, src_c):
            if not g.text(r, c):
                try:
                    v = f(g.text(r, src_c))
                except (IndexError, AttributeError):
                    v = None
                if v is not None:
                    changes.append((r, c, v))
            r += 1
        if changes:
            app.set_cells(changes)
            app.status["text"] = f"Flash fill: {len(changes)} cells"
            return len(changes)
    app.status["text"] = "Flash fill: no pattern found - type one or two more examples"
    return 0


# ---------- pivot table ----------

AGGS = {"Sum": sum, "Count": len, "Average": lambda v: sum(v) / len(v) if v else 0,
        "Min": lambda v: min(v) if v else 0, "Max": lambda v: max(v) if v else 0}


def make_pivot(rows, header, row_field, value_field, agg="Sum", col_field=None):
    """rows: table (first row = header). Returns a new table: row_field values down, col_field values
    across (or one value column), aggregated value_field."""
    idx = {h: i for i, h in enumerate(header)}
    ri, vi, ci = idx[row_field], idx[value_field], idx.get(col_field) if col_field else None
    groups, row_keys, col_keys = {}, [], []
    for row in rows:
        rk = row[ri] if ri < len(row) else ""
        ck = (row[ci] if ci < len(row) else "") if ci is not None else value_field
        if rk not in row_keys:
            row_keys.append(rk)
        if ck not in col_keys:
            col_keys.append(ck)
        v = row[vi] if vi < len(row) else ""
        n = parse_number(v)
        groups.setdefault((rk, ck), []).append(n if n is not None else (1 if agg == "Count" and v != "" else None))
    fn = AGGS[agg]

    def cell(vals):
        vals = [v for v in vals if v is not None]
        if agg == "Count":
            return str(len(vals))
        x = fn(vals)
        return str(int(x)) if float(x).is_integer() else f"{x:.10g}"
    title = f"{agg} of {value_field}"
    out = [[row_field] + ([str(k) for k in col_keys] if ci is not None else [title])]
    totals = [[] for _ in col_keys]
    for rk in sorted(row_keys, key=lambda k: (parse_number(k) is None, parse_number(k) or 0, str(k).casefold())):
        line = [rk]
        for j, ck in enumerate(col_keys):
            vals = groups.get((rk, ck), [])
            totals[j] += vals
            line.append(cell(vals) if vals else "")
        out.append(line)
    out.append(["Total"] + [cell(t) for t in totals])
    return out


class PivotDialog(tk.Toplevel):
    def __init__(self, app, rect):
        super().__init__(app)
        self.app, self.rect = app, rect
        g = app.table
        r1, c1, r2, c2 = rect
        self.header = [g.text(r1, c) or col_letters(c) for c in range(c1, c2 + 1)]
        self.rows = [[g.value(r, c) for c in range(c1, c2 + 1)] for r in range(r1 + 1, r2 + 1)]
        self.title("Pivot table")
        self.transient(app)
        f = ttk.Frame(self, padding=12)
        f.pack(fill="both", expand=True)
        self.boxes = {}
        fields = (("rows", "Rows", self.header), ("cols", "Columns (optional)", [""] + self.header),
                  ("values", "Values", self.header), ("agg", "Summarise by", list(AGGS)))
        for i, (key, label, values) in enumerate(fields):
            ttk.Label(f, text=label).grid(row=i, column=0, sticky="w", pady=2)
            cb = ttk.Combobox(f, values=values, state="readonly", width=24)
            cb.set(values[min(1, len(values) - 1)] if key == "values" else values[0])
            cb.grid(row=i, column=1, sticky="w", pady=2)
            self.boxes[key] = cb
        self.real = tk.BooleanVar(master=self, value=True)
        ttk.Checkbutton(f, text="Save as a real Excel pivot table (Excel refreshes it when opening)",
                        variable=self.real).grid(row=4, column=0, columnspan=2, sticky="w", pady=(6, 0))
        b = ttk.Frame(f)
        b.grid(row=5, column=0, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(b, text="Create", command=self.ok).pack(side="left", padx=4)
        ttk.Button(b, text="Cancel", command=self.destroy).pack(side="left")

    def ok(self):
        table = make_pivot(self.rows, self.header, self.boxes["rows"].get(), self.boxes["values"].get(),
                           self.boxes["agg"].get(), self.boxes["cols"].get() or None)
        real = self.real.get()
        source = (self.app.sheet,) + tuple(self.rect)
        spec = {"source": source, "rows": self.header.index(self.boxes["rows"].get()),
                "cols": self.header.index(self.boxes["cols"].get()) if self.boxes["cols"].get() else None,
                "values": self.header.index(self.boxes["values"].get()), "agg": self.boxes["agg"].get()}
        self.destroy()
        name = add_table_sheet(self.app, "Pivot", table)
        if real:
            self.app.pivots.append({**spec, "sheet": name, "saved": False})
        return name


def add_table_sheet(app, base, table, styles=None):
    """A new sheet after the current one holding table; styles: {(r, c): style} (default: summary look)."""
    import sheet_io as io
    n, name = 1, base
    while app.check_name(name):
        n += 1
        name = f"{base}{n}"
    s = io.Sheet([list(r) for r in table])
    if styles is not None:
        s.styles = dict(styles)
    else:
        s.styles = {(0, c): {"bold": True, "bg": "#dde5f3"} for c in range(len(table[0]))}
        s.styles.update({(len(table) - 1, c): {"bold": True} for c in range(len(table[-1]))})
    s.reset_structure()

    def mutate():
        items = list(app.sheets.items())
        items.insert(list(app.sheets).index(app.sheet) + 1, (name, s))
        app.sheets.clear()
        app.sheets.update(items)
    return name if app.book_change(mutate, name) else None


def pivot(app):
    return PivotDialog(app, _region(app))


# ---------- goal seek ----------

def goal_seek_solve(app, target, value, changing, tol=1e-7, max_iter=100):
    """Find the number for cell `changing` that makes formula cell `target` equal `value`.
    Cells are (r, c) on the current sheet. Returns the number or None."""
    g, eng = app.table, app.engine
    sheet = g.sheet
    name = app.sheet
    rows = sheet.rows
    r, c = changing
    while len(rows) <= r:
        rows.append([])
    while len(rows[r]) <= c:
        rows[r].append("")
    original = rows[r][c]

    def f(x):
        rows[r][c] = repr(float(x))
        eng.clear()
        v = eng.cell(name, *target)
        return (v - value) if isinstance(v, (int, float)) and not isinstance(v, bool) else None
    try:
        x0 = parse_number(original) or 0.0
        x1 = x0 + 1 if x0 == 0 else x0 * 1.1
        f0, f1 = f(x0), f(x1)
        for _ in range(max_iter):
            if f0 is None or f1 is None:
                return None
            if abs(f1) < tol:
                return x1
            if f1 == f0:
                x1 += 1e-3
                f1 = f(x1)
                continue
            x0, x1, f0 = x1, x1 - f1 * (x1 - x0) / (f1 - f0), f1
            f1 = f(x1)
        return x1 if f1 is not None and abs(f1) < 1e-4 else None
    finally:
        rows[r][c] = original
        eng.clear()


def goal_seek(app):
    from formula import parse_ref
    g = app.table
    here = f"{col_letters(g.cur[1])}{g.cur[0] + 1}"
    target = simpledialog.askstring("Goal seek", "Set cell (a formula):", initialvalue=here, parent=app)
    if not target:
        return None
    goal = simpledialog.askstring("Goal seek", "To value:", parent=app)
    change = simpledialog.askstring("Goal seek", "By changing cell:", parent=app)
    if goal is None or not change:
        return None
    try:
        _, tr, tc, _, _ = parse_ref("ref", target.strip().upper())
        _, cr, cc, _, _ = parse_ref("ref", change.strip().upper())
        want = float(goal)
    except (AttributeError, ValueError):
        messagebox.showwarning("Goal seek", "Type cells like B5 and a number.", parent=app)
        return None
    x = goal_seek_solve(app, (tr, tc), want, (cr, cc))
    if x is None:
        messagebox.showinfo("Goal seek", "No solution found.", parent=app)
        return None
    x = round(x, 10)
    app.set_cells([(cr, cc, str(int(x)) if float(x).is_integer() else f"{x:.10g}")])
    return x


# ---------- advanced filter ----------

def criterion_ok(value, crit):
    """One criteria cell (Excel style: '>5', '<>x', '=East', 'Ea*', 'East' = begins with) against a value."""
    import fnmatch
    crit = crit.strip()
    if not crit:
        return True
    for op in (">=", "<=", "<>", ">", "<", "="):
        if crit.startswith(op):
            want = crit[len(op):]
            a, b = parse_number(value), parse_number(want)
            if a is not None and b is not None:
                return {">=": a >= b, "<=": a <= b, "<>": a != b, ">": a > b, "<": a < b, "=": a == b}[op]
            t, w = value.casefold(), want.casefold()
            if op in ("=", "<>") and any(ch in w for ch in "*?"):
                hit = fnmatch.fnmatchcase(t, w)
                return hit if op == "=" else not hit
            return {">=": t >= w, "<=": t <= w, "<>": t != w, ">": t > w, "<": t < w, "=": t == w}[op]
    a, b = parse_number(value), parse_number(crit)
    if a is not None and b is not None:
        return a == b
    w = crit.casefold()
    return fnmatch.fnmatchcase(value.casefold(), w if any(ch in w for ch in "*?") else w + "*")


def advanced_filter(app, list_rect, crit_rect, copy_to=None):
    """Data → Advanced filter: the criteria range's first row names columns of the list; each row below is
    one alternative (cells in a row must all match). Hides the other rows, or copies the matching ones
    (with the header) to copy_to=(r, c). Returns the matching data rows."""
    g = app.table
    r1, c1, r2, c2 = list_rect
    head = [g.value(r1, c).strip().casefold() for c in range(c1, c2 + 1)]
    cr1, cc1, cr2, cc2 = crit_rect
    names = [g.value(cr1, c).strip().casefold() for c in range(cc1, cc2 + 1)]
    alternatives = []
    for r in range(cr1 + 1, cr2 + 1):
        conds = [(head.index(n), g.value(r, cc1 + k)) for k, n in enumerate(names)
                 if n in head and g.value(r, cc1 + k).strip()]
        alternatives.append(conds)
    alternatives = [a for a in alternatives if a] or [[]]
    matching = [r for r in range(r1 + 1, r2 + 1)
                if any(all(criterion_ok(g.value(r, c1 + j), crit) for j, crit in alt) for alt in alternatives)]
    if copy_to is None:
        changes = [("hr", r, 0, r not in matching) for r in range(r1 + 1, r2 + 1)]
        app.apply_mixed(changes)
    else:
        tr, tc = copy_to
        out = [[g.value(r1, c) for c in range(c1, c2 + 1)]] + [[g.value(r, c) for c in range(c1, c2 + 1)]
                                                                 for r in matching]
        app.set_cells([(tr + i, tc + j, v) for i, row in enumerate(out) for j, v in enumerate(row)])
    return matching


def advanced(app):
    from rules_dialog import parse_ranges, ref_text
    list_rect = _region(app)
    crit = simpledialog.askstring("Advanced filter", "Criteria range (header row + conditions), e.g. H1:I3:",
                                  parent=app)
    if not crit:
        return None
    crit_rect = parse_ranges(crit)
    if not crit_rect:
        messagebox.showwarning("Advanced filter", "Type a range like H1:I3.", parent=app)
        return None
    where = simpledialog.askstring("Advanced filter", f"List: {ref_text(list_rect)}\n\nCopy the results to (a cell), "
                                   "or leave empty to filter in place:", parent=app)
    copy_to = parse_ranges(where)[0][:2] if where and parse_ranges(where) else None
    return advanced_filter(app, list_rect, crit_rect[0], copy_to)
