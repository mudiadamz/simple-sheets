"""Tools → Compare with another workbook: every cell that differs (values and formulas), sheets only in one of
them; a list you can click through, or a report sheet."""
import os
import tkinter as tk
from tkinter import filedialog, ttk

from formula import col_letters


def diff_sheet(a, b):
    """Two sheets' rows -> [(r, c, text in a, text in b)] where they differ."""
    out = []
    for r in range(max(len(a), len(b))):
        ra = a[r] if r < len(a) else []
        rb = b[r] if r < len(b) else []
        for c in range(max(len(ra), len(rb))):
            x = ra[c] if c < len(ra) else ""
            y = rb[c] if c < len(rb) else ""
            if x != y and not _same_number(x, y):
                out.append((r, c, x, y))
    return out


def _same_number(x, y):
    from formatting import parse_number
    a, b = parse_number(x), parse_number(y)
    return a is not None and b is not None and a == b


def diff_books(mine, other):
    """{name: Sheet} x2 -> {"only_mine": [...], "only_other": [...], "cells": [(sheet, r, c, mine, other)]}."""
    cells = []
    for name, s in mine.items():
        if name in other:
            cells += [(name, *d) for d in diff_sheet(s.rows, other[name].rows)]
    return {"only_mine": [n for n in mine if n not in other], "only_other": [n for n in other if n not in mine],
            "cells": cells}


def compare_file(app, path):
    import sheet_io as io
    return diff_books(app.sheets, io.load(path))


class CompareWindow(tk.Toplevel):
    def __init__(self, app, path, result):
        super().__init__(app)
        self.app, self.path, self.result = app, path, result
        self.title(f"Compare with {os.path.basename(path)}")
        self.transient(app)
        f = ttk.Frame(self, padding=8)
        f.pack(fill="both", expand=True)
        n = len(result["cells"])
        parts = [f"{n} cell{'s' if n != 1 else ''} differ"]
        if result["only_mine"]:
            parts.append("only here: " + ", ".join(result["only_mine"]))
        if result["only_other"]:
            parts.append("only in the other file: " + ", ".join(result["only_other"]))
        self.summary = ttk.Label(f, text=" · ".join(parts), wraplength=640, justify="left")
        self.summary.pack(anchor="w")
        cols = ("sheet", "cell", "here", "other")
        self.tree = ttk.Treeview(f, columns=cols, show="headings", height=16)
        for c, title, w in zip(cols, ("Sheet", "Cell", "This workbook", os.path.basename(path)), (110, 60, 220, 220)):
            self.tree.heading(c, text=title)
            self.tree.column(c, width=w, stretch=c in ("here", "other"))
        for i, (sheet, r, c, x, y) in enumerate(result["cells"][:20000]):
            self.tree.insert("", "end", iid=str(i), values=(sheet, f"{col_letters(c)}{r + 1}", x, y))
        self.tree.pack(fill="both", expand=True, pady=6)
        self.tree.bind("<<TreeviewSelect>>", lambda e: self.go())
        b = ttk.Frame(f)
        b.pack(fill="x")
        ttk.Button(b, text="Report sheet", command=self.report).pack(side="left")
        ttk.Button(b, text="Close", command=self.destroy).pack(side="right")

    def go(self, i=None):
        sel = self.tree.selection()
        i = int(sel[0]) if i is None and sel else i
        if i is None:
            return None
        sheet, r, c, _x, _y = self.result["cells"][i]
        if self.app.sheet != sheet:
            self.app.show(sheet)
        self.app.table.goto(r, c)
        return sheet, r, c

    def report(self):
        import data_more
        rows = [["Sheet", "Cell", "This workbook", os.path.basename(self.path)]]
        rows += [[s, f"{col_letters(c)}{r + 1}", x, y] for s, r, c, x, y in self.result["cells"]]
        rows += [[s, "(sheet only here)", "", ""] for s in self.result["only_mine"]]
        rows += [[s, "(sheet only in the other file)", "", ""] for s in self.result["only_other"]]
        styles = {(0, c): {"bold": True, "bg": "#dde5f3"} for c in range(4)}
        self.destroy()
        name = data_more.add_table_sheet(self.app, "Differences", rows, styles)
        self.app.show(name)
        return name


def open_compare(app, path=None):
    from simple_sheets import OPEN_TYPES
    app.table.commit()
    path = path or filedialog.askopenfilename(parent=app, title="Compare with…", filetypes=OPEN_TYPES)
    if not path:
        return None
    try:
        result = compare_file(app, path)
    except Exception as e:
        from tkinter import messagebox
        messagebox.showerror("Compare", f"Couldn't open the file:\n{e}", parent=app)
        return None
    return CompareWindow(app, path, result)
