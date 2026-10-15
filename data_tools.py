"""Data tools: sort a range, and filter rows (autofilter with ▼ buttons in a header row)."""
import tkinter as tk
from tkinter import ttk

import formula
from grid import col_name
from theme import C

# ---------- sort ----------

def sort_key(v):
    """Excel order: numbers, then text (A-Z, case-insensitive), then FALSE/TRUE, then errors.
    Blanks are handled separately (always last)."""
    if isinstance(v, bool):
        return (2, int(v))
    if isinstance(v, (int, float)):
        return (0, float(v))
    if isinstance(v, formula.XLError):
        return (3, str(v))
    return (1, str(v).casefold())


def has_header(app, rect):
    """Guess: the first row is a header if it's all text and a column below it holds numbers,
    or if the first row is bold."""
    s, _g = app.current, app.table
    r1, c1, r2, c2 = rect
    if r2 <= r1:
        return False
    first = [app.engine.cell(app.sheet, r1, c) for c in range(c1, c2 + 1)]
    if any((s.styles.get((r1, c)) or {}).get("bold") for c in range(c1, c2 + 1)):
        return True
    if not all(isinstance(v, str) and v for v in first):
        return False
    below = [app.engine.cell(app.sheet, r, c) for r in range(r1 + 1, min(r2, r1 + 20) + 1)
             for c in range(c1, c2 + 1)]
    return any(isinstance(v, (int, float)) and not isinstance(v, bool) for v in below)


def sort_rect(app):
    """What a quick sort works on: the selection, or the data block around the active cell."""
    g = app.table
    return g.sel_bounded() if g.multi() else g.region(*g.cur)


def sort_range(app, rect, keys, header=None):
    """Sort rows of rect=(r1, c1, r2, c2). keys: [(column index, ascending)], first = most important.
    Values, formats (and formulas, with their references moved) travel with their row. One undo step."""
    g, s = app.table, app.current
    r1, c1, r2, c2 = rect
    if header is None:
        header = has_header(app, rect)
    top = r1 + 1 if header else r1
    if top >= r2:   # one row (or none) below the header: nothing to sort
        return False
    if any(m[0] <= r2 and m[2] >= top and m[1] <= c2 and m[3] >= c1 for m in s.merges):
        from tkinter import messagebox
        messagebox.showwarning("Sort", "Can't sort a range that contains merged cells.")
        return False
    app.engine.clear()
    rows = list(range(top, r2 + 1))

    def row_key(r):
        parts = []
        for key in keys:
            c, asc = key[0], key[1]
            order_list = [x.casefold() for x in key[2]] if len(key) > 2 and key[2] else None
            v = app.engine.cell(app.sheet, r, c)
            blank = v is None or v == ""
            k = sort_key(v) if not blank else (0, 0)
            if order_list is not None and not blank:   # a custom order (Mon, Tue… / Low, Medium, High)
                low = str(v).casefold()
                k = (0, order_list.index(low)) if low in order_list else (1, k)
            parts.append((blank, k, asc))
        return parts

    keyed = {r: row_key(r) for r in rows}
    order = list(rows)
    for i in reversed(range(len(keys))):   # stable sort, least important key first
        order.sort(key=lambda r: keyed[r][i][1], reverse=not keys[i][1])
        order.sort(key=lambda r: keyed[r][i][0])   # blanks last, whatever the direction
    changes = []
    for new_r, old_r in zip(rows, order):
        for c in range(c1, c2 + 1):
            t = g.text(old_r, c)
            if t.startswith("=") and new_r != old_r:
                t = formula.shift(t, new_r - old_r, 0)
            changes += [("v", new_r, c, t), ("s", new_r, c, s.styles.get((old_r, c)))]
    cur = g.cur
    app.apply_mixed(changes)
    g.select(r1, c1, r2, c2)
    if r1 <= cur[0] <= r2 and c1 <= cur[1] <= c2:   # keep the active cell (and so the sort column)
        g.cur = cur
        g.redraw()
        app.on_select()
    return True


def quick_sort(app, ascending=True):
    """Sort by the active cell's column (right-click 'Sort A → Z / Z → A')."""
    app.table.commit()
    rect = sort_rect(app)
    c = app.table.cur[1]
    return sort_range(app, rect, [(c, ascending)])


def custom_orders():
    """Sort orders besides A → Z: your fill lists (Settings) and day / month names. {label: [values]}"""
    import fill
    out = {}
    for lst in fill.custom_lists() + fill.CYCLES[:3]:
        label = ", ".join(lst[:3]) + ", …"
        out[label] = lst
    return out


class SortDialog(tk.Toplevel):
    """Sort range: up to three columns, each A → Z or Z → A, with a header-row option."""

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title("Sort range")
        self.transient(app)
        self.resizable(False, False)
        self.configure(background=C["bg"])
        self.rect = sort_rect(app)
        r1, c1, r2, c2 = self.rect
        f = ttk.Frame(self, padding=10)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text=f"Range {col_name(c1)}{r1 + 1}:{col_name(c2)}{r2 + 1}").grid(row=0, column=0,
                                                                                    columnspan=3, sticky="w")
        self.header = tk.BooleanVar(self, value=has_header(app, self.rect))
        ttk.Checkbutton(f, text="Data has a header row", variable=self.header,
                        command=self.fill_columns).grid(row=1, column=0, columnspan=3, sticky="w", pady=(4, 8))
        self.cols, self.orders, self.boxes = [], [], []
        for i in range(3):
            ttk.Label(f, text="Sort by" if i == 0 else "then by").grid(row=2 + i, column=0, sticky="w")
            col = tk.StringVar(self)
            box = ttk.Combobox(f, textvariable=col, state="readonly", width=22)
            box.grid(row=2 + i, column=1, padx=4, pady=2)
            order = tk.StringVar(self, value="A → Z")
            ttk.Combobox(f, textvariable=order, values=["A → Z", "Z → A"] + list(custom_orders()), state="readonly",
                         width=18).grid(row=2 + i, column=2)
            self.cols.append(col)
            self.orders.append(order)
            self.boxes.append(box)
        b = ttk.Frame(f)
        b.grid(row=6, column=0, columnspan=3, sticky="e", pady=(10, 0))
        ttk.Button(b, text="Sort", command=self.ok).pack(side="left", padx=2)
        ttk.Button(b, text="Cancel", command=self.destroy).pack(side="left")
        self.bind("<Escape>", lambda e: self.destroy())
        self.fill_columns()
        self.cols[0].set(self.labels[app.table.cur[1] - c1] if c1 <= app.table.cur[1] <= c2 else self.labels[0])

    def fill_columns(self):
        r1, c1, r2, c2 = self.rect
        g = self.app.table
        self.labels = []
        for c in range(c1, c2 + 1):
            head = g.value(r1, c) if self.header.get() else ""
            self.labels.append(f"Column {col_name(c)}" + (f" – {head}" if head else ""))
        for i, box in enumerate(self.boxes):
            box["values"] = ([] if i == 0 else ["(none)"]) + self.labels
            if i and not self.cols[i].get():
                self.cols[i].set("(none)")

    def ok(self):
        c1 = self.rect[1]
        keys = []
        lists = custom_orders()
        for col, order in zip(self.cols, self.orders):
            if col.get() in self.labels:
                keys.append((c1 + self.labels.index(col.get()), order.get() != "Z → A", lists.get(order.get())))
        if keys:
            sort_range(self.app, self.rect, keys, header=self.header.get())
        self.destroy()


# ---------- filter ----------
# sheet.filter = {"range": (r1, c1, r2, c2), "criteria": {col: [shown values kept]}}  (or None)
# Row r1 is the header (gets the ▼ buttons); rows hidden by the filter are in sheet.filtered_rows.

BLANK = "(Blanks)"


def shown(app, sheet_name, r, c):
    s = app.sheets[sheet_name]
    t = s.rows[r][c] if r < len(s.rows) and c < len(s.rows[r]) else ""
    if len(t) > 1 and t[0] == "=":
        return app.engine.display(s, r, c)
    return t


def recompute(app, sheet_name):
    """Which data rows the criteria hide."""
    s = app.sheets[sheet_name]
    s.filtered_rows = set()
    f = s.filter
    if not f:
        return
    r1, c1, r2, c2 = f["range"]
    tops = {}
    for c, keep in f["criteria"].items():       # top / bottom N: the threshold, worked out once
        if isinstance(keep, dict) and keep.get("cond") and keep["cond"][0] in ("top", "bottom"):
            nums = sorted(n for n in (formula.parse_number(shown(app, sheet_name, r, c)) for r in range(r1 + 1, r2 + 1))
                          if n is not None)
            k = max(int(formula.parse_number(str(keep["cond"][1])) or 10), 1)
            tops[c] = (nums[-k] if keep["cond"][0] == "top" else nums[k - 1]) if len(nums) >= k else None
    for r in range(r1 + 1, r2 + 1):
        for c, keep in f["criteria"].items():
            v = shown(app, sheet_name, r, c) or BLANK
            if not passes(s, r, c, v, keep, tops.get(c)):
                s.filtered_rows.add(r)
                break


COND_OPS = [("equals", "="), ("doesn't equal", "<>"), ("greater than", ">"), ("at least", ">="),
            ("less than", "<"), ("at most", "<="), ("contains", "contains"), ("doesn't contain", "!contains"),
            ("begins with", "begins"), ("ends with", "ends"), ("top N", "top"), ("bottom N", "bottom")]


def passes(sheet, r, c, v, keep, threshold=None):
    """Does the shown value v of cell (r, c) pass one column's criterion?"""
    if isinstance(keep, list):
        return v in keep
    if "color" in keep:
        return ((sheet.styles.get((r, c)) or {}).get("bg") or "") == (keep["color"] or "")
    op, want = keep["cond"]
    text = "" if v == BLANK else v
    if op in ("top", "bottom"):
        n = formula.parse_number(text)
        return n is not None and threshold is not None and (n >= threshold if op == "top" else n <= threshold)
    a, b = formula.parse_number(text), formula.parse_number(str(want))
    if op in ("=", "<>", ">", ">=", "<", "<=") and a is not None and b is not None:
        return {"=": a == b, "<>": a != b, ">": a > b, ">=": a >= b, "<": a < b, "<=": a <= b}[op]
    t, w = text.casefold(), str(want).casefold()
    return {"=": t == w, "<>": t != w, ">": t > w, ">=": t >= w, "<": t < w, "<=": t <= w,
            "contains": w in t, "!contains": w not in t, "begins": t.startswith(w), "ends": t.endswith(w)}[op]


def values_in(app, sheet_name, c):
    """Distinct shown values of a filter column (numbers sorted as numbers)."""
    r1, _, r2, _ = app.sheets[sheet_name].filter["range"]
    vals = {shown(app, sheet_name, r, c) or BLANK for r in range(r1 + 1, r2 + 1)}

    def k(v):
        n = formula.parse_number(v) if v != BLANK else None
        return (2, "") if v == BLANK else (0, n) if n is not None else (1, v.casefold())
    return sorted(vals, key=k)


def toggle_filter(app):
    """Create a filter on the selection / data block, or remove the sheet's filter."""
    g = app.table
    g.commit()
    s = app.current
    if s.filter:
        app.apply_mixed([("flt", 0, 0, None)])
        return
    r1, c1, r2, c2 = sort_rect(app)
    if r2 <= r1:
        r2 = r1 + 1
    app.apply_mixed([("flt", 0, 0, {"range": (r1, c1, r2, c2), "criteria": {}})])


def set_criteria(app, c, keep):
    """keep: list of shown values to keep for column c (None = no criteria on that column)."""
    f = app.current.filter
    import copy
    crit = copy.deepcopy(f["criteria"])
    if keep is None:
        crit.pop(c, None)
    else:
        crit[c] = copy.deepcopy(keep) if isinstance(keep, dict) else list(keep)
    app.apply_mixed([("flt", 0, 0, {**f, "range": tuple(f["range"]), "criteria": crit})])


class FilterPopup(tk.Toplevel):
    """▼ menu of a filter column: sort, then tick the values to show."""

    def __init__(self, app, c, x_root, y_root):
        super().__init__(app)
        self.app, self.c = app, c
        self.wm_overrideredirect(True)
        self.attributes("-topmost", True)
        self.configure(background=C["header_line"])
        f = ttk.Frame(self, padding=6)
        f.pack(padx=1, pady=1)
        ttk.Button(f, text="Sort A → Z", style="Toolbutton", command=lambda: self.sort(True)).pack(fill="x")
        ttk.Button(f, text="Sort Z → A", style="Toolbutton", command=lambda: self.sort(False)).pack(fill="x")
        ttk.Separator(f).pack(fill="x", pady=4)
        cond = ttk.Frame(f)                       # a condition: greater than 10, contains "x", top 5…
        cond.pack(fill="x")
        self.op = ttk.Combobox(cond, values=[k for k, _ in COND_OPS], state="readonly", width=13)
        self.op.set(COND_OPS[2][0])
        self.op.pack(side="left")
        self.want = ttk.Entry(cond, width=10)
        self.want.pack(side="left", padx=2)
        ttk.Button(cond, text="Apply", style="Toolbutton", command=self.apply_cond).pack(side="left")
        colors = sorted({(app.current.styles.get((r, c)) or {}).get("bg") or ""
                         for r in range(app.current.filter["range"][0] + 1, app.current.filter["range"][2] + 1)})
        if len(colors) > 1:                        # some cells are coloured: filter by fill colour
            bycol = ttk.Frame(f)
            bycol.pack(fill="x", pady=(4, 0))
            ttk.Label(bycol, text="By colour:").pack(side="left")
            for col in colors:
                sw = tk.Label(bycol, text="  " if col else "none", background=col or C["bg"], relief="solid", bd=1,
                              cursor="hand2")
                sw.pack(side="left", padx=2)
                sw.bind("<Button-1>", lambda e, col=col: self.by_color(col))
        ttk.Separator(f).pack(fill="x", pady=4)
        links = ttk.Frame(f)
        links.pack(fill="x")
        ttk.Button(links, text="Select all", style="Toolbutton", command=lambda: self.all(True)).pack(side="left")
        ttk.Button(links, text="Clear", style="Toolbutton", command=lambda: self.all(False)).pack(side="left")
        lst = ttk.Frame(f)
        lst.pack(fill="both", expand=True, pady=4)
        canvas = tk.Canvas(lst, height=min(220, 22 * 10), width=200, highlightthickness=0, background=C["bg"])
        bar = ttk.Scrollbar(lst, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window(0, 0, window=inner, anchor="nw")
        canvas.configure(yscrollcommand=bar.set)
        canvas.pack(side="left", fill="both", expand=True)
        bar.pack(side="right", fill="y")
        keep = app.current.filter["criteria"].get(c)
        keep = keep if isinstance(keep, list) else None
        self.vars = {}
        for v in values_in(app, app.sheet, c):
            var = tk.BooleanVar(self, value=keep is None or v in keep)
            ttk.Checkbutton(inner, text=v, variable=var).pack(anchor="w")
            self.vars[v] = var
        b = ttk.Frame(f)
        b.pack(fill="x")
        ttk.Button(b, text="OK", command=self.ok).pack(side="right")
        ttk.Button(b, text="Cancel", command=self.destroy).pack(side="right", padx=4)
        self.bind("<Escape>", lambda e: self.destroy())
        self.wm_geometry(f"+{x_root}+{y_root}")
        self.focus_force()

    def apply_cond(self):
        op = dict(COND_OPS)[self.op.get()]
        set_criteria(self.app, self.c, {"cond": [op, self.want.get().strip()]})
        self.destroy()

    def by_color(self, color):
        set_criteria(self.app, self.c, {"color": color})
        self.destroy()

    def all(self, on):
        for var in self.vars.values():
            var.set(on)

    def ok(self):
        keep = [v for v, var in self.vars.items() if var.get()]
        set_criteria(self.app, self.c, None if len(keep) == len(self.vars) else keep)
        self.destroy()

    def sort(self, ascending):
        r1, c1, r2, c2 = self.app.current.filter["range"]
        sort_range(self.app, (r1, c1, r2, c2), [(self.c, ascending)], header=True)
        self.destroy()
