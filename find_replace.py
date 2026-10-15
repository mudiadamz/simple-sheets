"""Find & replace (Ctrl+F / Ctrl+H): search values or formulas in the sheet or the whole workbook."""
import re
import tkinter as tk
from tkinter import ttk

from theme import C


def matcher(query, match_case=False, whole=False):
    """-> regex for the search text (plain text, not a pattern)."""
    body = re.escape(query)
    return re.compile(f"^{body}$" if whole else body, 0 if match_case else re.I)


def find_all(app, query, match_case=False, whole=False, in_formulas=False, workbook=False):
    """[(sheet name, r, c)] of matching cells, sheet by sheet (current sheet first), row by row."""
    if not query:
        return []
    rx = matcher(query, match_case, whole)
    names = list(app.sheets)
    order = names[names.index(app.sheet):] + names[:names.index(app.sheet)] if workbook else [app.sheet]
    hits = []
    for name in order:
        s = app.sheets[name]
        for r, row in enumerate(s.rows):
            for c, t in enumerate(row):
                if not t:
                    continue
                if not in_formulas and len(t) > 1 and t[0] == "=":
                    shown = app.engine.display(s, r, c)
                else:
                    shown = t
                if rx.search(shown):
                    hits.append((name, r, c))
    return hits


def replace_in(text, query, repl, match_case=False, whole=False):
    rx = matcher(query, match_case, whole)
    return rx.sub(lambda m: repl, text)


class FindDialog(tk.Toplevel):
    """Small non-modal window; stays on top of the app."""

    def __init__(self, app, replace=False):
        super().__init__(app)
        self.app = app
        self.title("Find and replace")
        self.transient(app)
        self.resizable(False, False)
        self.configure(background=C["bg"])
        f = ttk.Frame(self, padding=10)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text="Find").grid(row=0, column=0, sticky="w")
        self.find = ttk.Entry(f, width=34)
        self.find.grid(row=0, column=1, columnspan=3, sticky="we", pady=2)
        ttk.Label(f, text="Replace with").grid(row=1, column=0, sticky="w", padx=(0, 8))
        self.repl = ttk.Entry(f, width=34)
        self.repl.grid(row=1, column=1, columnspan=3, sticky="we", pady=2)
        self.match_case = tk.BooleanVar(self)
        self.whole = tk.BooleanVar(self)
        self.formulas = tk.BooleanVar(self)
        self.workbook = tk.BooleanVar(self)
        ttk.Checkbutton(f, text="Match case", variable=self.match_case, command=self.reset).grid(
            row=2, column=1, sticky="w")
        ttk.Checkbutton(f, text="Entire cell", variable=self.whole, command=self.reset).grid(
            row=2, column=2, sticky="w")
        ttk.Checkbutton(f, text="Search in formulas", variable=self.formulas, command=self.reset).grid(
            row=3, column=1, sticky="w")
        ttk.Checkbutton(f, text="All sheets", variable=self.workbook, command=self.reset).grid(
            row=3, column=2, sticky="w")
        buttons = ttk.Frame(f)
        buttons.grid(row=4, column=0, columnspan=4, sticky="e", pady=(8, 0))
        ttk.Button(buttons, text="Find all", command=self.find_all_list).pack(side="left", padx=2)
        ttk.Button(buttons, text="Find next", command=self.next).pack(side="left", padx=2)
        ttk.Button(buttons, text="Replace", command=self.replace_one).pack(side="left", padx=2)
        ttk.Button(buttons, text="Replace all", command=self.replace_all).pack(side="left", padx=2)
        ttk.Button(buttons, text="Close", command=self.destroy).pack(side="left", padx=2)
        self.status = ttk.Label(f, text="", foreground=C["muted"])
        self.status.grid(row=5, column=0, columnspan=4, sticky="w", pady=(6, 0))
        # "Find all": every match listed; click one to go there
        self.results = ttk.Treeview(f, columns=("sheet", "cell", "value"), show="headings", height=8)
        for col, w in (("sheet", 110), ("cell", 60), ("value", 220)):
            self.results.heading(col, text=col.capitalize())
            self.results.column(col, width=w)
        self.results.bind("<<TreeviewSelect>>", lambda e: self.go_result())
        self.results_shown = False
        self.results_frame_row = 6
        self.frame = f
        self.find.bind("<KeyRelease>", lambda e: self.reset(), add="+")
        self.find.bind("<Return>", lambda e: self.next())
        self.repl.bind("<Return>", lambda e: self.replace_one())
        self.bind("<Escape>", lambda e: self.destroy())
        self.hits, self.pos = None, -1
        (self.repl if replace else self.find).focus_set()

    def options(self):
        return dict(match_case=self.match_case.get(), whole=self.whole.get())

    def reset(self):
        self.hits, self.pos = None, -1
        self.status["text"] = ""

    def search(self):
        if self.hits is None:
            self.hits = find_all(self.app, self.find.get(), in_formulas=self.formulas.get(),
                                 workbook=self.workbook.get(), **self.options())
            # "next" starts after the active cell: later on this sheet, else the other sheets, else wrap
            here, cur = self.app.sheet, self.app.table.cur
            after = [i for i, (n, r, c) in enumerate(self.hits) if n == here and (r, c) > cur]
            elsewhere = [i for i, (n, r, c) in enumerate(self.hits) if n != here]
            start = after[0] if after else elsewhere[0] if elsewhere else 0
            self.pos = start - 1
        return self.hits

    def find_all_list(self):
        """List every match (sheet, cell, what it shows); picking a line goes to that cell."""
        from formula import col_letters
        self.reset()
        hits = self.search()
        self.results.delete(*self.results.get_children())
        for i, (name, r, c) in enumerate(hits):
            s = self.app.sheets[name]
            t = s.rows[r][c]
            shown = self.app.engine.display(s, r, c) if t.startswith("=") and len(t) > 1 else t
            self.results.insert("", "end", iid=str(i), values=(name, f"{col_letters(c)}{r + 1}", shown[:80]))
        if not self.results_shown:
            self.results.grid(row=self.results_frame_row, column=0, columnspan=4, sticky="nsew", pady=(6, 0))
            self.results_shown = True
        self.status["text"] = f"{len(hits)} found" if hits else "No matches"
        return hits

    def go_result(self):
        sel = self.results.selection()
        if not sel or not self.hits:
            return None
        self.pos = int(sel[0])
        name, r, c = self.hits[self.pos]
        if name != self.app.sheet:
            self.app.show(name)
            self.app.tabs.select(name)
        self.app.table.goto(r, c)
        return name, r, c

    def next(self):
        hits = self.search()
        if not hits:
            self.status["text"] = "No matches"
            return None
        self.pos = (self.pos + 1) % len(hits)
        name, r, c = hits[self.pos]
        if name != self.app.sheet:
            self.app.show(name)
            self.app.tabs.select(name)
        self.app.table.goto(r, c)
        self.status["text"] = f"{self.pos + 1} of {len(hits)}"
        return hits[self.pos]

    def replace_one(self):
        """Replace in the found cell (formulas: only when searching in formulas), then go on."""
        g = self.app.table
        hits = self.search()
        if self.pos < 0 or not hits:
            self.next()
            return
        name, r, c = hits[self.pos]
        if (name, r, c) == (self.app.sheet,) + g.cur:
            t = g.text(r, c)
            if self.formulas.get() or not t.startswith("="):
                new = replace_in(t, self.find.get(), self.repl.get(), **self.options())
                if new != t:
                    self.app.set_cells([(r, c, new)])
            self.hits = [h for h in hits if h != (name, r, c)]
            self.pos -= 1
        self.next()

    def replace_all(self):
        """Every match, every sheet searched, in one undo step."""
        hits = self.search()
        changes = []
        for name, r, c in hits:
            s = self.app.sheets[name]
            t = s.rows[r][c]
            if t.startswith("=") and not self.formulas.get():
                continue
            new = replace_in(t, self.find.get(), self.repl.get(), **self.options())
            if new != t:
                changes.append(("v", s, r, c, new))
        if changes:
            self.app.apply_cells(changes)
        self.status["text"] = f"Replaced {len(changes)} cell{'s' if len(changes) != 1 else ''}"
        self.reset_keep_status()
        return len(changes)

    def reset_keep_status(self):
        text = self.status["text"]
        self.reset()
        self.status["text"] = text


def open_dialog(app, replace=False):
    """One dialog per window; Ctrl+F / Ctrl+H focus the right field."""
    d = getattr(app, "find_dialog", None)
    if d is None or not d.winfo_exists():
        d = app.find_dialog = FindDialog(app, replace)
    else:
        d.deiconify()
        d.lift()
        (d.repl if replace else d.find).focus_set()
    g = app.table
    sel = g.value(*g.cur) if not g.multi() else ""
    if sel and not d.find.get():
        d.find.insert(0, sel)
        d.reset()
    return "break"
