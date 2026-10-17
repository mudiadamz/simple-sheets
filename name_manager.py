"""Name manager: list, add and delete defined names (named ranges) used in formulas like =SUM(prices)."""
import re
import tkinter as tk
from tkinter import messagebox, ttk

from simplecore.theme import C

import formula
from grid import col_name

VALID = re.compile(r"[A-Za-z_\\][\w.]*")


def selection_ref(app):
    """The current selection as an absolute reference with its sheet: 'Sheet1'!$A$1:$B$5."""
    g = app.table
    r1, c1, r2, c2 = g.sel_bounded() if g.multi() else g.span(*g.cur)
    a = f"${col_name(c1)}${r1 + 1}"
    b = f"${col_name(c2)}${r2 + 1}"
    return f"{formula.quote_sheet(app.sheet)}!{a}" + ("" if (r1, c1) == (r2, c2) else f":{b}")


def check_name(app, name, refers):
    """Error text or None."""
    if not VALID.fullmatch(name) or re.fullmatch(r"[A-Za-z]{1,3}\d+", name) or name.upper() in ("TRUE", "FALSE"):
        return "Names start with a letter or _, use letters, digits, _ or . and can't look like a cell (A1)."
    if name.upper() in formula.FUNCS:
        return f"“{name}” is a function name."
    toks = [(k, t) for k, t in formula.lex(refers.lstrip("=")) if k != "ws"]
    if len(toks) != 1 or toks[0][0] not in ("ref", "cols", "rows"):
        return "“Refers to” must be one cell or range, like Sheet1!$A$1:$B$5."
    return None


class NameManager(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title("Named ranges")
        self.transient(app)
        self.configure(background=C["bg"])
        f = ttk.Frame(self, padding=10)
        f.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(f, columns=("ref",), height=8, selectmode="browse")
        self.tree.heading("#0", text="Name")
        self.tree.heading("ref", text="Refers to")
        self.tree.column("#0", width=150)
        self.tree.column("ref", width=240)
        self.tree.grid(row=0, column=0, columnspan=4, sticky="nsew")
        ttk.Label(f, text="Name").grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.name = ttk.Entry(f, width=18)
        self.name.grid(row=1, column=1, sticky="w", pady=(8, 0))
        ttk.Label(f, text="Refers to").grid(row=2, column=0, sticky="w")
        self.refers = ttk.Entry(f, width=34)
        self.refers.grid(row=2, column=1, columnspan=3, sticky="we")
        self.refers.insert(0, selection_ref(app))
        b = ttk.Frame(f)
        b.grid(row=3, column=0, columnspan=4, sticky="e", pady=(8, 0))
        ttk.Button(b, text="Add / update", command=self.add).pack(side="left", padx=2)
        ttk.Button(b, text="Delete", command=self.delete).pack(side="left", padx=2)
        ttk.Button(b, text="Close", command=self.destroy).pack(side="left", padx=2)
        self.tree.bind("<<TreeviewSelect>>", self.picked)
        self.bind("<Escape>", lambda e: self.destroy())
        self.refresh()
        self.name.focus_set()

    def refresh(self):
        self.tree.delete(*self.tree.get_children())
        for n, t in sorted(self.app.names.items(), key=lambda x: x[0].lower()):
            self.tree.insert("", "end", iid=n, text=n, values=(t,))

    def picked(self, _=None):
        sel = self.tree.selection()
        if sel:
            self.name.delete(0, "end")
            self.name.insert(0, sel[0])
            self.refers.delete(0, "end")
            self.refers.insert(0, self.app.names[sel[0]])

    def add(self):
        name, refers = self.name.get().strip(), self.refers.get().strip().lstrip("=")
        err = check_name(self.app, name, refers)
        if err:
            messagebox.showwarning("Named ranges", err, parent=self)
            return False
        names = {k: v for k, v in self.app.names.items() if k.lower() != name.lower()}
        names[name] = refers
        self.app.set_names(names)
        self.refresh()
        return True

    def delete(self):
        name = self.name.get().strip()
        names = {k: v for k, v in self.app.names.items() if k.lower() != name.lower()}
        if names != self.app.names:
            self.app.set_names(names)
        self.refresh()
