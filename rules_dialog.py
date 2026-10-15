"""Create / delete conditional formatting rules and data validation for the selection (Data menu)."""
import re
import tkinter as tk
from tkinter import messagebox, ttk

from formula import col_index, col_letters

CF_TYPES = [
    ("Cell value", "cellIs"), ("Text contains", "containsText"), ("Text doesn't contain", "notContainsText"),
    ("Begins with", "beginsWith"), ("Ends with", "endsWith"), ("Formula is true", "expression"),
    ("Blank", "containsBlanks"), ("Not blank", "notContainsBlanks"), ("Duplicate values", "duplicateValues"),
    ("Unique values", "uniqueValues"), ("Top N", "top"), ("Bottom N", "bottom"),
    ("Above average", "above"), ("Below average", "below"),
    ("Colour scale (2 colours)", "scale2"), ("Colour scale (3 colours)", "scale3"), ("Data bar", "dataBar"),
]
OPS = [("greater than", "greaterThan"), ("less than", "lessThan"), ("between", "between"),
       ("not between", "notBetween"), ("equal to", "equal"), ("not equal to", "notEqual"),
       ("at least", "greaterThanOrEqual"), ("at most", "lessThanOrEqual")]
STYLES = [
    ("Light red fill, dark red text", {"bg": "#ffc7ce", "fg": "#9c0006"}),
    ("Yellow fill, dark yellow text", {"bg": "#ffeb9c", "fg": "#9c5700"}),
    ("Green fill, dark green text", {"bg": "#c6efce", "fg": "#006100"}),
    ("Light red fill", {"bg": "#ffc7ce"}),
    ("Red text", {"fg": "#9c0006"}),
    ("Bold", {"bold": True}),
]
KEEP = "(keep its current look)"
SCALES = {"scale2": ["#f8696b", "#63be7b"], "scale3": ["#f8696b", "#ffeb84", "#63be7b"]}
DV_TYPES = [("Any value", None), ("List of items", "list"), ("Whole number", "whole"), ("Decimal", "decimal"),
            ("Date", "date"), ("Text length", "textLength"), ("Custom formula", "custom")]


def ref_text(rect):
    r1, c1, r2, c2 = rect
    a = f"{col_letters(c1)}{r1 + 1}"
    return a if (r1, c1) == (r2, c2) else f"{a}:{col_letters(c2)}{r2 + 1}"


def parse_ranges(text):
    """'A1:B5 D2' -> [(r1, c1, r2, c2)] or None."""
    out = []
    for part in re.split(r"[\s,;]+", text.strip()):
        if not part:
            continue
        m = re.fullmatch(r"\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?", part)
        if not m:
            return None
        r1, c1 = int(m.group(2)) - 1, col_index(m.group(1))
        r2, c2 = (int(m.group(4)) - 1, col_index(m.group(3))) if m.group(3) else (r1, c1)
        out.append((min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)))
    return out or None


def describe(d):
    t = d["type"]
    if t == "cellIs":
        words = dict((v, k) for k, v in OPS).get(d.get("op"), d.get("op"))
        return f"Cell value {words} " + " and ".join(str(f) for f in d.get("f", []) if f not in (None, ""))
    if t in ("containsText", "notContainsText", "beginsWith", "endsWith"):
        return f"{dict((v, k) for k, v in CF_TYPES).get(t, t)} “{d.get('text')}”"
    if t == "expression":
        return f"Formula: ={d['f'][0] if d.get('f') else ''}"
    if t == "top10":
        return f"{'Bottom' if d.get('bottom') else 'Top'} {d.get('rank')}{'%' if d.get('percent') else ''}"
    if t == "aboveAverage":
        return "Above average" if d.get("above", True) else "Below average"
    if t == "colorScale":
        return "Colour scale"
    if t == "dataBar":
        return "Data bar"
    return dict((v, k) for k, v in CF_TYPES).get(t, t)


def subtract(rect, cut):
    """Parts of rect outside cut (up to 4 rectangles)."""
    r1, c1, r2, c2 = rect
    a1, b1, a2, b2 = cut
    if a1 > r2 or a2 < r1 or b1 > c2 or b2 < c1:
        return [rect]
    out = []
    if r1 < a1:
        out.append((r1, c1, a1 - 1, c2))
    if r2 > a2:
        out.append((a2 + 1, c1, r2, c2))
    top, bottom = max(r1, a1), min(r2, a2)
    if c1 < b1:
        out.append((top, c1, bottom, b1 - 1))
    if c2 > b2:
        out.append((top, b2 + 1, bottom, c2))
    return out


def open_dialog(app, kind):
    return (ConditionalDialog if kind == "conditional" else ValidationDialog)(app)


# ---------- conditional formatting ----------

def make_cf(kind, ranges, op=None, v1="", v2="", style=None):
    """Build an in-app rule dict (see rules.read_conditional) - positions are the sheet's own."""
    d = {"ranges": ranges, "anchor": ranges[0][:2], "model": True, "type": kind, "op": None, "f": [],
         "priority": 0, "stop": False, "style": dict(style or {}), "text": None, "rank": None, "percent": False,
         "bottom": False, "above": True, "equal": False}
    if kind == "cellIs":
        d["op"], d["f"] = op, [v.lstrip("=") for v in (v1, v2) if v != ""]
        if op not in ("between", "notBetween"):
            d["f"] = d["f"][:1]
    elif kind in ("containsText", "notContainsText", "beginsWith", "endsWith"):
        d["text"] = v1
    elif kind == "expression":
        d["f"] = [v1.lstrip("=")]
    elif kind in ("top", "bottom"):
        d["type"], d["rank"], d["bottom"] = "top10", int(float(v1 or 10)), kind == "bottom"
        d["percent"] = v2.strip().endswith("%") if v2 else False
    elif kind in ("above", "below"):
        d["type"], d["above"] = "aboveAverage", kind == "above"
    elif kind in ("scale2", "scale3"):
        cols = SCALES[kind]
        pts = [("min", None, cols[0])] + ([("percentile", 50, cols[1])] if kind == "scale3" else []) + \
              [("max", None, cols[-1])]
        d["type"], d["scale"], d["style"] = "colorScale", pts, {}
    elif kind == "dataBar":
        d["bar"], d["style"] = ([("min", None), ("max", None)], "#638ec6"), {}
    return d


class ConditionalDialog(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.sheet = app.current
        self.title("Conditional formatting")
        self.transient(app)
        f = ttk.Frame(self, padding=10)
        f.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(f, columns=("range", "rule", "format"), show="headings", height=6)
        for col, w in (("range", 110), ("rule", 230), ("format", 160)):
            self.tree.heading(col, text=col.capitalize())
            self.tree.column(col, width=w)
        self.tree.pack(fill="both", expand=True)
        row = ttk.Frame(f)
        row.pack(fill="x", pady=(4, 8))
        ttk.Button(row, text="Delete rule", command=self.delete).pack(side="left")
        ttk.Button(row, text="Edit rule", command=self.edit_selected).pack(side="left", padx=4)
        ttk.Button(row, text="▲", width=3, command=lambda: self.move(-1)).pack(side="left")
        ttk.Button(row, text="▼", width=3, command=lambda: self.move(1)).pack(side="left")
        ttk.Label(row, text="(rules higher in the list win)", foreground="gray").pack(side="left", padx=6)
        self.editing = None          # index of the rule being edited ("Add rule" then saves it)
        self.kept_style = None       # an edited rule's own look when it isn't one of STYLES
        add = ttk.LabelFrame(f, text="New rule", padding=8)
        add.pack(fill="x")
        ttk.Label(add, text="Apply to").grid(row=0, column=0, sticky="w")
        self.range = tk.StringVar(master=self, value=" ".join(ref_text(r) for r in app.table.cells_rects()))
        ttk.Entry(add, textvariable=self.range, width=28).grid(row=0, column=1, columnspan=2, sticky="w")
        ttk.Label(add, text="Format cells if").grid(row=1, column=0, sticky="w")
        self.kind = ttk.Combobox(add, values=[k for k, _ in CF_TYPES], state="readonly", width=26)
        self.kind.current(0)
        self.kind.grid(row=1, column=1, columnspan=2, sticky="w")
        self.op = ttk.Combobox(add, values=[k for k, _ in OPS], state="readonly", width=14)
        self.op.current(0)
        self.op.grid(row=2, column=0, sticky="w")
        self.v1 = tk.StringVar(master=self)
        self.v2 = tk.StringVar(master=self)
        ttk.Entry(add, textvariable=self.v1, width=14).grid(row=2, column=1, sticky="w")
        ttk.Entry(add, textvariable=self.v2, width=14).grid(row=2, column=2, sticky="w")
        ttk.Label(add, text="Style").grid(row=3, column=0, sticky="w")
        self.style = ttk.Combobox(add, values=[k for k, _ in STYLES], state="readonly", width=30)
        self.style.current(0)
        self.style.grid(row=3, column=1, columnspan=2, sticky="w")
        self.add_btn = ttk.Button(add, text="Add rule", command=self.add)
        self.add_btn.grid(row=4, column=2, sticky="e", pady=(6, 0))
        ttk.Button(f, text="Close", command=self.destroy).pack(anchor="e", pady=(8, 0))
        self.bind("<Escape>", lambda e: self.destroy())
        self.fill()

    def fill(self):
        self.tree.delete(*self.tree.get_children())
        self.rows = list(self.sheet.cond)
        for i, d in enumerate(self.rows):
            st = d.get("style") or {}
            fmt = ", ".join(f"{k} {v}" if v is not True else k for k, v in st.items()) or "-"
            self.tree.insert("", "end", iid=str(i), values=(" ".join(ref_text(r) for r in d["ranges"]),
                                                          describe(d), fmt))

    def add(self, kind=None, op=None, v1=None, v2=None, style=None, ranges=None):
        ranges = ranges or parse_ranges(self.range.get())
        if not ranges:
            messagebox.showwarning("Conditional formatting", "Type a range like A2:A20.", parent=self)
            return False
        kind = kind or dict(CF_TYPES)[self.kind.get()]
        op = op or dict(OPS)[self.op.get()]
        v1 = self.v1.get() if v1 is None else v1
        v2 = self.v2.get() if v2 is None else v2
        if kind in ("cellIs", "expression", "containsText", "notContainsText", "beginsWith", "endsWith") and not v1:
            messagebox.showwarning("Conditional formatting", "Type the value to compare with.", parent=self)
            return False
        if style is None:
            style = self.kept_style if self.style.get() == KEEP and self.kept_style is not None \
                else dict(STYLES).get(self.style.get(), {})
        d = make_cf(kind, ranges, op, v1, v2, style)
        rules = list(self.sheet.cond)
        if self.editing is not None and self.editing < len(rules):
            rules[self.editing] = d                       # edited in place (keeps its place in the order)
        else:
            rules.insert(0, d)                            # new rules win
        self.app.apply_cells([("cf", self.sheet, 0, 0, rules)])
        self.stop_editing()
        self.fill()
        return True

    def edit_selected(self):
        """Load the selected rule into the form; "Save changes" then replaces it."""
        sel = self.tree.selection()
        if not sel:
            return False
        i = int(sel[0])
        d = self.rows[i]
        self.editing = i
        self.range.set(" ".join(ref_text(r) for r in d["ranges"]))
        t = d["type"]
        kind = {"top10": "bottom" if d.get("bottom") else "top",
                "aboveAverage": "above" if d.get("above", True) else "below",
                "colorScale": "scale3" if len(d.get("scale") or []) == 3 else "scale2"}.get(t, t)
        self.kind.set(next((k for k, v in CF_TYPES if v == kind), CF_TYPES[0][0]))
        self.op.set(next((k for k, v in OPS if v == d.get("op")), OPS[0][0]))
        f = [str(x) for x in d.get("f") or []]
        if t in ("containsText", "notContainsText", "beginsWith", "endsWith"):
            f = [d.get("text") or ""]
        elif t == "top10":
            f = [str(d.get("rank") or 10), "%" if d.get("percent") else ""]
        self.v1.set(f[0] if f else "")
        self.v2.set(f[1] if len(f) > 1 else "")
        st = d.get("style") or {}
        preset = next((k for k, v in STYLES if v == st), None)
        self.kept_style = None if preset else st
        self.style.configure(values=[k for k, _ in STYLES] + ([KEEP] if not preset else []))
        self.style.set(preset or KEEP)
        self.add_btn.configure(text="Save changes")
        return True

    def stop_editing(self):
        self.editing, self.kept_style = None, None
        self.style.configure(values=[k for k, _ in STYLES])
        if self.style.get() not in dict(STYLES):
            self.style.current(0)
        self.add_btn.configure(text="Add rule")

    def move(self, step):
        """▲ / ▼: rules higher in the list are applied first (they win where formats clash)."""
        sel = self.tree.selection()
        if not sel:
            return False
        i = int(sel[0])
        j = i + step
        if not 0 <= j < len(self.rows):
            return False
        rules = list(self.rows)
        rules[i], rules[j] = rules[j], rules[i]
        self.app.apply_cells([("cf", self.sheet, 0, 0, rules)])
        self.fill()
        self.tree.selection_set(str(j))
        return True

    def delete(self):
        sel = self.tree.selection()
        if not sel:
            return
        gone = {int(i) for i in sel}
        self.app.apply_cells([("cf", self.sheet, 0, 0, [d for i, d in enumerate(self.rows) if i not in gone])])
        self.fill()


# ---------- data validation ----------

class ValidationDialog(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.sheet = app.current
        g = app.table
        self.rects = g.cells_rects()
        self.title("Data validation")
        self.transient(app)
        self.resizable(False, False)
        r, c = g.span(*g.cur)[:2]
        cur = app.rules.validation_at(self.sheet, r, c) or {}
        f = ttk.Frame(self, padding=12)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text="Cells: " + " ".join(ref_text(x) for x in self.rects)).grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Label(f, text="Allow").grid(row=1, column=0, sticky="w")
        self.kind = ttk.Combobox(f, values=[k for k, _ in DV_TYPES], state="readonly", width=20)
        self.kind.set(next((k for k, v in DV_TYPES if v == cur.get("type")), DV_TYPES[0][0]))
        self.kind.grid(row=1, column=1, columnspan=2, sticky="w")
        ttk.Label(f, text="Condition").grid(row=2, column=0, sticky="w")
        self.op = ttk.Combobox(f, values=[k for k, _ in OPS], state="readonly", width=20)
        self.op.set(next((k for k, v in OPS if v == cur.get("op")), "between"))
        self.op.grid(row=2, column=1, columnspan=2, sticky="w")
        ttk.Label(f, text="Value / list / formula").grid(row=3, column=0, sticky="w")
        f1 = cur.get("f1") or ""
        if cur.get("type") == "list" and f1.startswith('"'):
            f1 = f1.strip('"')
        self.v1 = tk.StringVar(master=self, value=f1)
        self.v2 = tk.StringVar(master=self, value=cur.get("f2") or "")
        ttk.Entry(f, textvariable=self.v1, width=24).grid(row=3, column=1, sticky="w")
        ttk.Entry(f, textvariable=self.v2, width=12).grid(row=3, column=2, sticky="w")
        ttk.Label(f, text="(list: a,b,c or a range like =$F$1:$F$5)", foreground="gray").grid(
            row=4, column=1, columnspan=2, sticky="w")
        self.arrow = tk.BooleanVar(master=self, value=cur.get("arrow", True))
        ttk.Checkbutton(f, text="Show drop-down arrow", variable=self.arrow).grid(row=5, column=1, columnspan=2, sticky="w")
        ttk.Label(f, text="On invalid input").grid(row=6, column=0, sticky="w")
        self.style = ttk.Combobox(f, values=["stop", "warning", "information"], state="readonly", width=12)
        self.style.set(cur.get("style", "stop"))
        self.style.grid(row=6, column=1, sticky="w")
        ttk.Label(f, text="Message").grid(row=7, column=0, sticky="w")
        self.msg = tk.StringVar(master=self, value=cur.get("error", ""))
        ttk.Entry(f, textvariable=self.msg, width=36).grid(row=7, column=1, columnspan=2, sticky="w")
        b = ttk.Frame(f)
        b.grid(row=8, column=0, columnspan=3, sticky="e", pady=(10, 0))
        ttk.Button(b, text="Clear", command=self.clear).pack(side="left", padx=4)
        ttk.Button(b, text="OK", command=self.ok).pack(side="left", padx=4)
        ttk.Button(b, text="Cancel", command=self.destroy).pack(side="left")
        self.bind("<Escape>", lambda e: self.destroy())

    def without_selection(self):
        """The sheet's rules with the selected cells cut out (rules from the file only when the sheet's
        layout still matches the file)."""
        out, same = [], not self.sheet.structure_changed()
        for d in self.sheet.validations:
            if not d.get("model") and not same:
                out.append(d)
                continue
            ranges = list(d["ranges"])
            for cut in self.rects:
                ranges = [p for r in ranges for p in subtract(r, cut)]
            if ranges == list(d["ranges"]):
                out.append(d)
            elif ranges:
                nd = dict(d, ranges=ranges, model=True)
                nd.pop("id", None)
                out.append(nd)
        return out

    def ok(self, kind=None):
        kind = kind if kind is not None else dict(DV_TYPES)[self.kind.get()]
        rules = self.without_selection()
        if kind:
            f1, f2 = self.v1.get().strip(), self.v2.get().strip()
            if not f1:
                messagebox.showwarning("Data validation", "Type the list items or a value.", parent=self)
                return False
            if kind == "list":
                f1 = f1[1:] if f1.startswith("=") else '"' + f1.strip('"') + '"'
            else:
                f1, f2 = f1.lstrip("="), f2.lstrip("=") or None
            rules.insert(0, {"ranges": list(self.rects), "model": True, "type": kind, "op": dict(OPS)[self.op.get()],
                             "f1": f1, "f2": f2 if kind not in ("list", "custom") else None, "arrow": self.arrow.get(),
                             "show_error": True, "style": self.style.get(), "title": "", "error": self.msg.get()})
        self.app.apply_cells([("dv", self.sheet, 0, 0, rules)])
        self.destroy()
        return True

    def clear(self):
        self.app.apply_cells([("dv", self.sheet, 0, 0, self.without_selection())])
        self.destroy()
        return True
