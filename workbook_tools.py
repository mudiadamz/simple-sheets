"""Workbook / sheet commands: view options (gridlines, headings, formulas, full screen), tab colours,
duplicate sheet, copy a sheet to a new workbook, sheet protection, document properties, calculation."""
import tkinter as tk
from tkinter import colorchooser, filedialog, messagebox, simpledialog, ttk

import sheet_io as io
import sheet_props
import structure

TAB_COLORS = [("Red", "#e06666"), ("Orange", "#f6b26b"), ("Yellow", "#ffd966"), ("Green", "#93c47d"),
              ("Teal", "#76a5af"), ("Blue", "#6fa8dc"), ("Purple", "#8e7cc3"), ("Pink", "#c27ba0"),
              ("Grey", "#999999")]


# ---------- view ----------

def toggle_sheet_setting(app, key):
    """Gridlines / headings on or off for the current sheet (saved in the file)."""
    app.table.commit()
    on = app.current.props.get(key) is not False
    app.apply_mixed([("sp", 0, key, None if not on else False)])


def toggle_formulas(app):
    app.show_formulas = not getattr(app, "show_formulas", False)
    app.formulas_var.set(app.show_formulas)
    app.table.redraw()


def toggle_full_screen(app):
    full = not bool(app.attributes("-fullscreen"))
    app.attributes("-fullscreen", full)
    app.full_var.set(full)
    return "break"


# ---------- sheets ----------

def set_tab_color(app, name, color):
    app.apply_cells([("sp", app.sheets[name], 0, "tab", color)])


def pick_tab_color(app, name):
    c = colorchooser.askcolor(color=app.sheets[name].props.get("tab"), parent=app, title="Tab colour")[1]
    if c:
        set_tab_color(app, name, c.lower())


def copy_sheet(src, dst_name=None):
    """A deep copy of a Sheet (cells, styles, layout, settings), as a new sheet not yet in the file."""
    snap = structure.snapshot({"s": src})["s"]
    s = io.Sheet([])
    structure.restore({"s": s}, {"s": snap})
    s.props = {k: v for k, v in src.props.items() if k != "page"}
    s.outline_rows, s.outline_cols = dict(src.outline_rows), dict(src.outline_cols)
    same = not src.structure_changed()        # file positions = sheet positions: rules can be copied as they are
    s.validations = [dict(d, model=True) for d in src.validations if same or d.get("model")]
    s.cond = [dict(d, model=True) for d in src.cond if same or d.get("model")]
    for d in s.cond + s.validations:
        d.pop("id", None)
    s.rowmap, s.rowtail, s.colmap, s.coltail = [], 0, [], 0
    s.disk_name = None
    s.hidden = False
    return s


def duplicate_sheet(app, name):
    base = f"{name} (2)"
    n = 2
    while app.check_name(base):
        n += 1
        base = f"{name} ({n})"
    new = copy_sheet(app.sheets[name])

    def mutate():
        items = list(app.sheets.items())
        at = [k for k, _ in items].index(name) + 1
        items.insert(at, (base, new))
        app.sheets.clear()
        app.sheets.update(items)
    app.book_change(mutate, base)
    return base


def copy_to_new_workbook(app, name, dst=None):
    """Save one sheet (values, formatting, layout) as a new .xlsx file."""
    app.table.commit()
    if dst is None:
        dst = filedialog.asksaveasfilename(filetypes=[("Excel workbook", "*.xlsx")], defaultextension=".xlsx",
                                           initialfile=f"{name}.xlsx", parent=app)
        if not dst:
            return None
    s = copy_sheet(app.sheets[name])
    try:
        io.write_xlsx(dst, {name: s}, {})
    except Exception as e:
        messagebox.showerror("Copy sheet", str(e), parent=app)
        return None
    app.status["text"] = f"“{name}” copied to {dst}"
    return dst


# ---------- protection ----------

def is_locked(app, sheet, r, c):
    """True when sheet protection forbids changing cell (r, c)."""
    if not sheet.props.get("protect"):
        return False
    st = sheet.styles.get((r, c)) or {}
    return not st.get("unlocked")


def protect_sheet(app, password=None):
    """Protect the current sheet; only cells marked 'unlocked' stay editable."""
    app.table.commit()
    if password is None:
        password = simpledialog.askstring("Protect sheet", "Password (optional):", show="•", parent=app)
        if password is None:
            return False
    prot = {"password": sheet_props.legacy_hash(password) if password else ""}
    app.apply_mixed([("sp", 0, "protect", prot)])
    return True


def unprotect_sheet(app, password=None):
    prot = app.current.props.get("protect")
    if not prot:
        return True
    if prot.get("password") or prot.get("hash"):
        if password is None:
            password = simpledialog.askstring("Unprotect sheet", "Password:", show="•", parent=app)
            if password is None:
                return False
        if not sheet_props.check_password(prot, password):
            messagebox.showerror("Unprotect sheet", "The password isn't right.", parent=app)
            return False
    app.apply_mixed([("sp", 0, "protect", None)])
    return True


def toggle_locked(app):
    """Format → mark the selection as unlocked (editable when the sheet is protected) or locked again."""
    g = app.table
    r, c = g.span(*g.cur)[:2]
    unlocked = bool((g.sheet.styles.get((r, c)) or {}).get("unlocked"))
    app.format_cells(lambda st: {**st, "unlocked": not unlocked})


# ---------- workbook properties / calculation ----------

class PropertiesDialog(tk.Toplevel):
    FIELDS = [("title", "Title"), ("subject", "Subject"), ("creator", "Author"), ("keywords", "Keywords"),
              ("category", "Category"), ("description", "Comments")]

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title("Properties")
        self.transient(app)
        self.resizable(False, False)
        f = ttk.Frame(self, padding=12)
        f.pack(fill="both", expand=True)
        self.vars = {}
        for i, (key, label) in enumerate(self.FIELDS):
            ttk.Label(f, text=label).grid(row=i, column=0, sticky="w", pady=2)
            v = tk.StringVar(master=self, value=app.book_props.get(key, ""))
            ttk.Entry(f, textvariable=v, width=44).grid(row=i, column=1, sticky="ew", pady=2)
            self.vars[key] = v
        b = ttk.Frame(f)
        b.grid(row=len(self.FIELDS), column=0, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(b, text="OK", command=self.ok).pack(side="left", padx=4)
        ttk.Button(b, text="Cancel", command=self.destroy).pack(side="left")
        self.bind("<Return>", lambda e: self.ok())
        self.bind("<Escape>", lambda e: self.destroy())

    def ok(self):
        self.app.apply_cells([("wp", None, 0, k, v.get().strip() or None) for k, v in self.vars.items()])
        self.destroy()


class CalculationDialog(tk.Toplevel):
    """Automatic / manual calculation, iterative calculation for circular references."""

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title("Calculation")
        self.transient(app)
        self.resizable(False, False)
        p = app.book_props
        f = ttk.Frame(self, padding=12)
        f.pack(fill="both", expand=True)
        self.mode = tk.StringVar(master=self, value=p.get("calc", "auto"))
        ttk.Radiobutton(f, text="Automatic", value="auto", variable=self.mode).grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(f, text="Manual (F9 recalculates)", value="manual", variable=self.mode).grid(row=1, column=0, sticky="w")
        self.iterate = tk.BooleanVar(master=self, value=bool(p.get("iterate")))
        ttk.Checkbutton(f, text="Iterative calculation (allow circular references)",
                        variable=self.iterate).grid(row=2, column=0, columnspan=2, sticky="w", pady=(10, 2))
        ttk.Label(f, text="Maximum iterations").grid(row=3, column=0, sticky="w")
        self.count = tk.StringVar(master=self, value=str(p.get("iterate_count", 100)))
        ttk.Entry(f, textvariable=self.count, width=8).grid(row=3, column=1, sticky="w")
        ttk.Label(f, text="Maximum change").grid(row=4, column=0, sticky="w")
        self.delta = tk.StringVar(master=self, value=str(p.get("iterate_delta", 0.001)))
        ttk.Entry(f, textvariable=self.delta, width=8).grid(row=4, column=1, sticky="w")
        b = ttk.Frame(f)
        b.grid(row=5, column=0, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(b, text="OK", command=self.ok).pack(side="left", padx=4)
        ttk.Button(b, text="Cancel", command=self.destroy).pack(side="left")

    def ok(self):
        try:
            count, delta = max(1, min(int(self.count.get()), 32767)), abs(float(self.delta.get()))
        except ValueError:
            messagebox.showwarning("Calculation", "Iterations must be a whole number and the change a number.",
                                   parent=self)
            return
        it = self.iterate.get()
        self.app.apply_cells([("wp", None, 0, "calc", "manual" if self.mode.get() == "manual" else None),
                              ("wp", None, 0, "iterate", True if it else None),
                              ("wp", None, 0, "iterate_count", count if it else None),
                              ("wp", None, 0, "iterate_delta", delta if it else None)])
        self.destroy()


def recalculate(app):
    """F9: calculate every formula again (the only way values update in manual mode)."""
    app.engine.clear()
    app.engine.rescan()
    app.table.data_changed()
    app.status["text"] = "Recalculated"
    return "break"
