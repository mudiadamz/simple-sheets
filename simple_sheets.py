"""Simple Sheets - a tiny spreadsheet viewer/editor for .xlsx / .xls / .ods / .csv files."""
import os
import subprocess
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import app_menu
import fill as fill_series
import formula
import menus
import recent
import settings
import sheet_io as io
import structure
import theme
from format_bar import FormatBar, Tooltip
from formatting import DEFAULT_SIZE, describe, parse_number
from formula import Engine
from grid import ALL, Grid, col_name
from sheet_tabs import SheetTabs

OPEN_TYPES = [("Spreadsheets", "*.xlsx *.xlsm *.xls *.ods *.csv *.tsv"), ("All files", "*.*")]
SAVE_TYPES = [("Excel workbook", "*.xlsx"), ("CSV (current sheet)", "*.csv")]
STATS_LIMIT = 200_000  # cells scanned for the Sum/Count status
ZOOMS = [0.5, 0.67, 0.75, 0.8, 0.9, 1.0, 1.1, 1.25, 1.5, 1.75, 2.0, 2.5, 3.0]
BAD_NAME_CHARS = set('[]:*?/\\')
NEWLINE_MARK = "↵"   # line breaks as shown in the one-line formula bar


def resource(*parts):
    """Path of a bundled file, both when run from source and from the built exe."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, *parts)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Simple Sheets")
        try:
            self.iconbitmap(default=resource("assets", "simple_sheets.ico"))
        except tk.TclError:
            pass
        self.scale = self.winfo_fpixels("1i") / 96   # 1.0 at 100% display scaling
        self.geometry(f"{int(1000 * self.scale)}x{int(600 * self.scale)}")   # size when un-maximized
        try:
            self.state("zoomed")    # open maximized (Windows)
        except tk.TclError:
            self.attributes("-zoomed", True)   # Linux
        self.path = None
        self.sheets = {}           # name -> io.Sheet, in tab order
        self.sheet = None          # current sheet name
        self.disk_sheets = []      # sheet names in the file on disk, in order
        self.base = {}             # (kind, Sheet, r, c) -> value on disk, for every touched cell
        self.history, self.hpos = [], 0   # [[(kind, Sheet or None, r, c, old, new), ...]], undo position
        self.select_after = None   # sheet to show after a sheet add/rename/delete
        self.formula_home = None   # (Sheet, r, c) of a formula being typed while another sheet is shown

        theme.base_styles(ttk.Style(self), self.scale)
        self.dark_var = tk.BooleanVar(self, value=bool(settings.get("dark", False)))

        # one toolbar row: ☰ menu, undo/redo, formatting
        top = ttk.Frame(self, padding=(4, 4, 0, 0))
        top.pack(fill="x")
        app_menu.build(self, top)
        self.fmtbar = FormatBar(top, self, self.scale)
        self.fmtbar.pack(side="left", fill="x", expand=True)

        # formula bar: cell ref | value (editable) | format of the active cell
        fbar = ttk.Frame(self, padding=(4, 0, 4, 4))
        fbar.pack(fill="x")
        self.ref = ttk.Label(fbar, width=12, anchor="center", relief="groove")
        self.ref.pack(side="left", ipady=2)
        self.value = ttk.Entry(fbar)
        self.value.pack(side="left", fill="x", expand=True, padx=4)
        self.value.bind("<Return>", self.value_commit)
        self.value.bind("<KeyRelease>", lambda e: self.table.show_refs(self.value.get()), add="+")
        self.value.bind("<FocusIn>", lambda e: self.table.show_refs(self.value.get()), add="+")
        self.value.bind("<FocusOut>", lambda e: self.table.show_refs(""), add="+")
        self.value.bind("<Escape>", self.value_cancel)
        self.fmt_info = ttk.Label(fbar, foreground=theme.C["muted"])
        self.fmt_info.pack(side="right")

        # bottom: sheet tabs + status
        bottom = ttk.Frame(self)
        bottom.pack(side="bottom", fill="x")
        self.tabs = SheetTabs(bottom, self)
        self.tabs.pack(side="left", fill="x", expand=True)
        self.zoom_label = ttk.Label(bottom, text="100%", width=5, anchor="e", cursor="hand2")
        self.zoom_label.pack(side="right", padx=(0, 8))
        self.zoom_label.bind("<Button-1>", lambda e: self.zoom_to(1.0))
        Tooltip(self.zoom_label, "Zoom: Ctrl + / Ctrl - / Ctrl+wheel. Click for 100%")
        ttk.Separator(bottom, orient="vertical").pack(side="right", fill="y", padx=4)
        self.status = ttk.Label(bottom, text="", padding=(8, 0))
        self.status.pack(side="right")

        self.engine = Engine()
        self.table = Grid(self, self, self.scale)
        self.table.pack(fill="both", expand=True)
        for k, key in (("b", "bold"), ("i", "italic"), ("u", "underline")):
            self.table.cv.bind(f"<Control-{k}>", lambda e, key=key: self.fmtbar.shortcut(key))
        self.table.cv.bind("<Control-d>", lambda e: self.fill_copy("down"))
        self.table.cv.bind("<Control-r>", lambda e: self.fill_copy("right"))

        self.bind("<Control-n>", lambda e: self.new_window())
        self.bind("<Control-o>", lambda e: self.open_dialog())
        self.bind("<Control-s>", lambda e: self.save())
        self.bind("<Control-S>", lambda e: self.save_as())  # Ctrl+Shift+S
        self.bind("<Control-z>", lambda e: self.undo())
        self.bind("<Control-y>", lambda e: self.redo())
        self.bind("<Control-Z>", lambda e: self.redo())  # Ctrl+Shift+Z
        for k in ("plus", "equal", "KP_Add"):
            self.bind(f"<Control-{k}>", lambda e: self.zoom_step(1))
        for k in ("minus", "KP_Subtract"):
            self.bind(f"<Control-{k}>", lambda e: self.zoom_step(-1))
        self.bind("<Control-0>", lambda e: self.zoom_to(1.0))
        self.bind("<Control-KP_0>", lambda e: self.zoom_to(1.0))
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.new_document()
        self.set_theme(self.dark_var.get())

    def toggle_dark(self):
        self.set_theme(self.dark_var.get())
        settings.set("dark", self.dark_var.get())

    def set_theme(self, dark):
        """Light / dark for the whole window."""
        self.dark_var.set(dark)
        theme.apply(self, dark)
        self.fmt_info.configure(foreground=theme.C["muted"])
        self.fmtbar.refresh_icons()
        self.__dict__.pop("menu_icons", None)   # right-click menu icons get redrawn in the new ink
        self.table.redraw()

    # ---------- model ----------
    # Changes and the "on disk" baseline refer to Sheet objects, so renaming a sheet keeps them valid.

    @property
    def current(self):
        return self.sheets[self.sheet]

    def name_of(self, sheet):
        return next((n for n, s in self.sheets.items() if s is sheet), None)

    def get(self, sheet, r, c):
        rows = sheet.rows
        return rows[r][c] if r < len(rows) and c < len(rows[r]) else ""

    def put(self, sheet, r, c, val):
        rows = sheet.rows
        while len(rows) <= r:
            rows.append([])
        row = rows[r]
        while len(row) <= c:
            row.append("")
        row[c] = val

    # every change is (kind, sheet, r, c): "v" = cell text, "s" = cell style dict,
    # "x" = rows/columns moved (whole-book snapshot), "b" = sheets added/renamed/deleted
    # "hr" / "hc" = row r / column c hidden (True/False)
    def read(self, kind, sheet, r, c):
        if kind == "m":                    # "m" = the sheet's merged ranges (whole list)
            return list(sheet.merges)
        if kind == "w":                    # "w" = width of column c (Excel characters; None = default)
            return sheet.widths.get(c)
        if kind == "hr":
            return r in sheet.hidden_rows
        if kind == "hc":
            return c in sheet.hidden_cols
        return self.get(sheet, r, c) if kind == "v" else sheet.styles.get((r, c))

    def write_cell(self, kind, sheet, r, c, val):
        if kind == "m":
            sheet.merges = list(val)
        elif kind == "w":
            if val is None:
                sheet.widths.pop(c, None)
            else:
                sheet.widths[c] = val
        elif kind in ("hr", "hc"):
            target, i = (sheet.hidden_rows, r) if kind == "hr" else (sheet.hidden_cols, c)
            (target.add if val else target.discard)(i)
        elif kind == "x":
            structure.restore(self.sheets, val)
        elif kind == "b":
            self.restore_book(val)
        elif kind == "v":
            self.put(sheet, r, c, val)
        elif val:
            sheet.styles[(r, c)] = val
        else:
            sheet.styles.pop((r, c), None)

    def apply(self, kind, changes):
        """Apply [(r, c, new), ...] on the current sheet as one undo step."""
        self.apply_mixed([(kind, r, c, val) for r, c, val in changes])

    def apply_mixed(self, changes):
        """Apply [(kind, r, c, new), ...] (values and styles) as one undo step."""
        s, group = self.current, []
        for kind, r, c, val in changes:
            old = self.read(kind, s, r, c)
            if val != old:
                self.base.setdefault((kind, s, r, c), old)
                group.append((kind, s, r, c, old, val))
                self.write_cell(kind, s, r, c, val)
        if group:
            self.push(group)

    def push(self, group):
        del self.history[self.hpos:]
        self.history.append(group)
        self.hpos += 1
        self.after_change(group)

    def restructure(self, axis, at, n):
        """Insert (n > 0) or delete (n < 0) rows/columns at `at` on the current sheet."""
        self.table.commit()
        if len(self.current.rows) >= io.MAX_ROWS:
            messagebox.showwarning("Too big", f"Only the first {io.MAX_ROWS:,} rows are loaded, so rows "
                                              "and columns can't be inserted or deleted in this sheet.")
            return
        before = structure.snapshot(self.sheets)
        for home, r, c in structure.change(self.sheets, self.sheet, axis, at, n):
            if home != self.sheet:   # formulas on other sheets: saved as normal cell edits
                self.base.setdefault(("v", self.sheets[home], r, c), before[home]["rows"][r][c])
        r, c = self.table.cur
        if axis == "row" and r >= at:
            r = max(r + n, at) if n < 0 else r + n
        if axis == "col" and c >= at:
            c = max(c + n, at) if n < 0 else c + n
        self.push([("x", self.current, r, c, before, structure.snapshot(self.sheets))])

    def shift_cells(self, direction):
        """Insert cells (shift right/down) or delete cells (shift left/up) for the selection."""
        g = self.table
        g.commit()
        r1, c1, r2, c2 = g.sel_cells()
        self.apply_mixed(structure.shift_cells(g.sheet, (r1, c1, r2, c2), direction))
        g.select(r1, c1, r2, c2)

    def set_cells(self, changes):
        self.apply("v", changes)

    def set_width(self, c, chars):
        """Column width (Excel characters) as an undoable, saved change."""
        self.apply_mixed([("w", 0, c, chars)])

    def fill(self, src, direction, count, copy=False):
        """Fill handle: continue the block src=(r1, c1, r2, c2) by `count` cells
        'down' / 'up' / 'right' / 'left'. copy=True repeats instead of counting up."""
        g = self.table
        g.commit()
        r1, c1, r2, c2 = src
        vertical = direction in ("down", "up")
        if direction == "up":
            count = min(count, r1)
        elif direction == "left":
            count = min(count, c1)
        if count <= 0:
            return
        changes = []
        for line in (range(c1, c2 + 1) if vertical else range(r1, r2 + 1)):
            cells = [(r, line) for r in range(r1, r2 + 1)] if vertical else [(line, c) for c in range(c1, c2 + 1)]
            if direction in ("up", "left"):
                cells.reverse()            # continue the run backwards
            texts = [g.text(*p) for p in cells]
            for k, (i, text) in enumerate(fill_series.extend(texts, count, copy)):
                sr, sc = cells[i]
                tr, tc = cells[-1]
                if direction == "down":
                    tr += k + 1
                elif direction == "up":
                    tr -= k + 1
                elif direction == "right":
                    tc += k + 1
                else:
                    tc -= k + 1
                if text is None:           # copy, moving formula references
                    text = texts[i]
                    if text.startswith("="):
                        text = formula.shift(text, tr - sr, tc - sc)
                changes += [("v", tr, tc, text), ("s", tr, tc, g.sheet.styles.get((sr, sc)))]
        self.apply_mixed(changes)
        r1, r2 = (r1, r2 + count) if direction == "down" else (r1 - count, r2) if direction == "up" else (r1, r2)
        c1, c2 = (c1, c2 + count) if direction == "right" else (c1 - count, c2) if direction == "left" else (c1, c2)
        g.select(r1, c1, r2, c2)

    def fill_to_neighbour(self):
        """Double-click on the fill handle: fill down as far as the column next to it goes."""
        g = self.table
        r1, c1, r2, c2 = g.sel_bounded() if g.multi() else g.span(*g.cur)
        for c in (c1 - 1, c2 + 1):
            if c < 0:
                continue
            end = r2
            while g.text(end + 1, c) != "":
                end += 1
            if end > r2:
                self.fill((r1, c1, r2, c2), "down", end - r2)
                return

    def fill_copy(self, axis):
        """Ctrl+D / Ctrl+R: copy the selection's first row / column over the rest
        (a single cell copies the cell above / left of it)."""
        g = self.table
        g.commit()
        r1, c1, r2, c2 = g.sel_bounded()
        if axis == "down":
            if r1 == r2:
                if r1 == 0:
                    return
                self.fill((r1 - 1, c1, r1 - 1, c2), "down", 1, copy=True)
            else:
                self.fill((r1, c1, r1, c2), "down", r2 - r1, copy=True)
        else:
            if c1 == c2:
                if c1 == 0:
                    return
                self.fill((r1, c1 - 1, r2, c1 - 1), "right", 1, copy=True)
            else:
                self.fill((r1, c1, r2, c1), "right", c2 - c1, copy=True)
        return "break"

    def merge(self, center=False):
        """Merge the selection into one cell (only the first value is kept, like Excel)."""
        g = self.table
        g.commit()
        r1, c1, r2, c2 = g.sel_bounded()
        s = self.current
        while True:   # grow to swallow merges that stick out of the selection
            hits = [m for m in s.merges if m[0] <= r2 and m[2] >= r1 and m[1] <= c2 and m[3] >= c1]
            grown = (min([r1] + [m[0] for m in hits]), min([c1] + [m[1] for m in hits]),
                     max([r2] + [m[2] for m in hits]), max([c2] + [m[3] for m in hits]))
            if grown == (r1, c1, r2, c2):
                break
            r1, c1, r2, c2 = grown
        if (r1, c1) == (r2, c2):
            return False
        if (r2 - r1 + 1) * (c2 - c1 + 1) > 100_000:
            messagebox.showwarning("Merge cells", "Select a smaller range to merge.")
            return False
        filled = [(r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1) if g.text(r, c) != ""]
        if len(filled) > 1 and not messagebox.askokcancel(
                "Merge cells", "Merging keeps only the upper-left value and clears the others."):
            return False
        changes = [("m", 0, 0, [m for m in s.merges if m not in hits] + [(r1, c1, r2, c2)])]
        if filled:
            keep = g.text(*filled[0])
            changes += [("v", r, c, "") for r, c in filled] + [("v", r1, c1, keep)]
        if center:
            st = dict(s.styles.get((r1, c1)) or {})
            changes.append(("s", r1, c1, {**st, "align": "center"}))
        self.apply_mixed(changes)
        g.select(r1, c1, r1, c1)
        return True

    def unmerge(self):
        """Split every merged block touching the selection."""
        g = self.table
        g.commit()
        r1, c1, r2, c2 = g.sel_bounded()
        s = self.current
        keep = [m for m in s.merges if not (m[0] <= r2 and m[2] >= r1 and m[1] <= c2 and m[3] >= c1)]
        if keep != s.merges:
            self.apply_mixed([("m", 0, 0, keep)])
            return True
        return False

    def hide(self, axis, hide=True):
        """Hide the selected rows/columns, or unhide the hidden ones inside the selection."""
        g = self.table
        g.commit()
        r1, c1, r2, c2 = g.sel()
        lo, hi, hidden = (r1, r2, self.current.hidden_rows) if axis == "row" else (c1, c2, self.current.hidden_cols)
        if hide:
            if hi >= ALL:
                return
            idx = range(lo, hi + 1)
        else:
            idx = sorted(i for i in hidden if lo <= i <= hi)
        kind = "hr" if axis == "row" else "hc"
        self.apply_mixed([(kind, i, 0, hide) if axis == "row" else (kind, 0, i, hide) for i in idx])
        if not hide:
            g.redraw()

    def format_cells(self, fn):
        """Restyle the selection: fn(style dict) -> new style dict."""
        self.table.commit()
        if self.table.sheet:
            self.format_rect(self.table.sel_cells(), lambda st, r, c: fn(st))

    def format_rect(self, rect, fn):
        """Restyle a block: fn(style dict, r, c) -> new style dict."""
        g = self.table
        r1, c1, r2, c2 = rect
        if (r2 - r1 + 1) * (c2 - c1 + 1) > 500_000:
            messagebox.showwarning("Too many cells", "Select a smaller range to format.")
            return
        changes = []
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                if g.span(r, c)[:2] != (r, c):
                    continue   # merged-away cell: the block's top-left holds the format
                st = fn(dict(g.sheet.styles.get((r, c)) or {}), r, c)
                st = {k: v for k, v in st.items() if v and not (k == "size" and v == DEFAULT_SIZE)}
                changes.append((r, c, st or None))
        self.apply("s", changes)

    def after_change(self, group):
        self.engine.clear()   # recalculate formulas on next draw
        kinds = {g[0] for g in group}
        if "b" in kinds:
            self.after_book_change()
        elif "x" in kinds:
            r, c = group[0][2], group[0][3]
            self.table.reload()
            self.table.goto(r, c)
        elif "s" in kinds:
            self.table.style_changed({g[2] for g in group})
        if kinds & {"hr", "hc", "m", "w"}:
            self.table.reload()
            self.table.fix_cursor()
        self.table.data_changed()
        self.refresh()

    def changes(self, kind):
        """{(sheet name, r, c): current value} for cells that differ from the file."""
        names = {id(s): n for n, s in self.sheets.items()}
        return {(names[id(k[1])],) + k[2:]: self.read(*k) for k, v in self.base.items()
                if k[0] == kind and id(k[1]) in names and self.read(*k) != v}

    def book_changed(self):
        return [s.disk_name for s in self.sheets.values()] != self.disk_sheets or \
            any(s.disk_name != n or s.hidden != s.disk_hidden for n, s in self.sheets.items())

    def visible_sheets(self):
        return [n for n, s in self.sheets.items() if not s.hidden]

    def dirty(self):
        live = {id(s) for s in self.sheets.values()}
        return self.book_changed() or any(s.structure_changed() for s in self.sheets.values()) or \
            any(self.read(*k) != v for k, v in self.base.items() if id(k[1]) in live)

    def refresh(self):
        self.undo_btn["state"] = "normal" if self.hpos > 0 else "disabled"
        self.redo_btn["state"] = "normal" if self.hpos < len(self.history) else "disabled"
        name = os.path.basename(self.path) if self.path else "Untitled"
        dirty = self.dirty()
        self.title(f"Simple Sheets - {'*' if dirty else ''}{name}")
        self.save_btn["state"] = "normal" if dirty else "disabled"
        self.on_select()

    # ---------- sheets: add / rename / delete ----------

    def book_state(self, select=None):
        """(sheet order, contents snapshot, sheet to show) - one side of a sheet add/rename/delete."""
        return list(self.sheets.items()), structure.snapshot(self.sheets), select or self.sheet

    def restore_book(self, state):
        order, snap, self.select_after = state
        self.sheets.clear()
        self.sheets.update(order)
        structure.restore(self.sheets, snap)

    def book_change(self, mutate, select):
        """Run mutate() (changes self.sheets) as one undo step, then show `select`."""
        self.table.commit()
        before = self.book_state()
        mutate()
        self.select_after = select
        self.push([("b", None, 0, 0, before, self.book_state(select))])

    def after_book_change(self):
        cur = self.table.sheet
        name = self.select_after if self.select_after in self.sheets else self.name_of(cur)
        self.select_after = None
        if name is None or self.sheets[name].hidden:
            name = self.visible_sheets()[0]
        self.sheet = name
        self.engine.set_book(self.sheets)
        self.tabs.rebuild()
        if self.table.sheet is not self.current:
            self.table.load(self.current)

    def fix_formulas(self, fn):
        """Rewrite formulas everywhere; changed cells are saved as normal edits."""
        for s in self.sheets.values():
            for r, row in enumerate(s.rows):
                for c, t in enumerate(row):
                    if t.startswith("=") and len(t) > 1:
                        new = fn(t)
                        if new != t:
                            self.base.setdefault(("v", s, r, c), t)
                            row[c] = new

    def check_name(self, name, old=None):
        """Error text, or None if `name` is a valid new sheet name."""
        if not name:
            return "The name can't be empty."
        if len(name) > 31:
            return "Use at most 31 characters."
        if BAD_NAME_CHARS & set(name) or name.startswith("'") or name.endswith("'"):
            return "A sheet name can't contain  [ ] : * ? / \\  or start/end with '."
        if any(n.lower() == name.lower() for n in self.sheets if n != old):
            return f"There is already a sheet named “{name}”."
        return None

    def add_sheet(self):
        n = len(self.sheets) + 1
        while self.check_name(f"Sheet{n}"):
            n += 1
        name = f"Sheet{n}"

        def mutate():
            items = list(self.sheets.items())
            at = list(self.sheets).index(self.sheet) + 1 if self.sheet in self.sheets else len(items)
            items.insert(at, (name, io.Sheet([])))
            self.sheets.clear()
            self.sheets.update(items)
        self.book_change(mutate, name)

    def rename_sheet(self, old, new):
        err = self.check_name(new, old)
        if err:
            messagebox.showwarning("Rename sheet", err)
            return False

        def mutate():
            items = [(new if n == old else n, s) for n, s in self.sheets.items()]
            self.sheets.clear()
            self.sheets.update(items)
            self.fix_formulas(lambda t: formula.rename_sheet(t, old, new))
        self.book_change(mutate, new)
        return True

    def move_sheet(self, name, index):
        """Move a sheet to position `index` among the visible tabs (hidden sheets keep their slots)."""
        def mutate():
            order = list(self.sheets)
            slots = [i for i, n in enumerate(order) if not self.sheets[n].hidden]
            vis = [order[i] for i in slots if order[i] != name]
            vis.insert(index, name)
            for i, n in zip(slots, vis):
                order[i] = n
            items = [(n, self.sheets[n]) for n in order]
            self.sheets.clear()
            self.sheets.update(items)
        self.book_change(mutate, name)

    def neighbour(self, name):
        """Visible sheet next to `name` (after it, else before)."""
        vis = self.visible_sheets()
        i = vis.index(name)
        return vis[i + 1] if i + 1 < len(vis) else vis[i - 1]

    def hide_sheet(self, name, hide=True):
        if hide and self.visible_sheets() == [name]:
            messagebox.showwarning("Hide sheet", "At least one sheet must stay visible.")
            return False
        select = self.neighbour(name) if hide else name

        def mutate():
            self.sheets[name].hidden = hide
        self.book_change(mutate, select)
        return True

    def delete_sheet(self, name):
        if self.visible_sheets() == [name]:
            messagebox.showwarning("Delete sheet", "A workbook needs at least one visible sheet.")
            return False
        if any(any(row) for row in self.sheets[name].rows) and not messagebox.askokcancel(
                "Delete sheet", f"Delete “{name}” and its data? (Undo brings it back.)"):
            return False
        neighbour = self.neighbour(name)

        def mutate():
            del self.sheets[name]
            self.fix_formulas(lambda t: formula.rename_sheet(t, name, None))
        self.book_change(mutate, neighbour)
        return True

    # ---------- undo / redo ----------

    def replay(self, group, undo):
        sheet = group[0][1]
        name = self.name_of(sheet) if sheet is not None else None
        if name and name != self.sheet:
            self.show(name)
            self.tabs.select(name)
        for kind, s, r, c, old, new in (reversed(group) if undo else group):
            self.write_cell(kind, s, r, c, old if undo else new)
        if group[0][0] in ("v", "s"):
            rs, cs = [g[2] for g in group], [g[3] for g in group]
            self.table.select(min(rs), min(cs), max(rs), max(cs))
        self.after_change(group)

    def undo(self):
        if self.formula_home is not None:
            self.end_cross_sheet(save=False)
        self.table.commit()
        if self.hpos > 0:
            self.hpos -= 1
            self.replay(self.history[self.hpos], undo=True)

    def redo(self):
        if self.formula_home is not None:
            self.end_cross_sheet(save=False)
        self.table.commit()
        if self.hpos < len(self.history):
            self.hpos += 1
            self.replay(self.history[self.hpos - 1], undo=False)

    # ---------- selection info ----------

    def on_select(self):
        g = self.table
        if not g.sheet:
            return
        r, c = g.span(*g.cur)[:2]
        r1, c1, r2, c2 = g.sel()
        if not g.multi():
            ref = f"{col_name(c)}{r + 1}"
        elif r2 >= ALL and c2 >= ALL:
            ref = "All"
        elif r2 >= ALL:
            ref = f"{col_name(c1)}:{col_name(c2)}"
        elif c2 >= ALL:
            ref = f"{r1 + 1}:{r2 + 1}"
        else:
            ref = f"{col_name(c1)}{r1 + 1}:{col_name(c2)}{r2 + 1}"
        self.ref["text"] = ref
        if self.focus_get() is not self.value and self.formula_home is None:
            self.value.delete(0, "end")
            self.value.insert(0, g.text(r, c).replace("\n", NEWLINE_MARK))
        self.fmt_info["text"] = "Format: " + describe(g.sheet.styles.get((r, c)))
        self.fmtbar.sync(g.sheet.styles.get((r, c)))
        self.fmtbar.merged.set((r, c) in g.merge_at)
        # Excel-style stats for multi-cell selections
        if g.multi():
            dr1, dc1, dr2, dc2 = g.sel_data()
            count, nums, total = 0, 0, 0.0
            if (dr2 - dr1 + 1) * (dc2 - dc1 + 1) <= STATS_LIMIT:
                for rr in range(dr1, dr2 + 1):
                    for cc in range(dc1, dc2 + 1):
                        t = g.value(rr, cc)
                        if t:
                            count += 1
                            n = parse_number(t)
                            if n is not None:
                                nums, total = nums + 1, total + n
            parts = [f"Count: {count:,}"]
            if nums:
                parts += [f"Sum: {total:,.10g}", f"Average: {total / nums:,.10g}"]
            self.status["text"] = "    ".join(parts)
        else:
            rows = len(g.sheet.rows)
            note = f" (first {io.MAX_ROWS:,} only)" if rows >= io.MAX_ROWS else ""
            self.status["text"] = f"{rows:,} rows{note}"

    def value_commit(self, _=None):
        g = self.table
        g.point = None
        if self.formula_home is not None:
            self.end_cross_sheet(save=True)
            return "break"
        if g.sheet:
            r, c = g.span(*g.cur)[:2]
            self.set_cells([(r, c, self.value.get().replace(NEWLINE_MARK, "\n"))])
            g.cv.focus_set()
            g.move(1, 0)
        return "break"

    def value_cancel(self, _=None):
        if self.formula_home is not None:
            self.end_cross_sheet(save=False)
        else:
            self.table.cv.focus_set()
            self.on_select()
        return "break"

    def zoom_step(self, d):
        z = self.table.zoom
        steps = [s for s in ZOOMS if (s > z + 0.001 if d > 0 else s < z - 0.001)]
        if steps:
            self.zoom_to(steps[0] if d > 0 else steps[-1])
        return "break"

    def zoom_to(self, z):
        self.table.set_zoom(z)
        self.zoom_label["text"] = f"{round(self.table.zoom * 100)}%"
        return "break"

    def context_menu(self, x, y):
        menus.show(self, x, y)

    # ---------- files ----------

    def load_book(self, path, sheets):
        """Show a workbook (sheets as on disk) in this window."""
        self.path, self.sheets = path, sheets
        self.formula_home = None
        for name, s in sheets.items():
            s.disk_name, s.disk_hidden = name, s.hidden
        if all(s.hidden for s in sheets.values()):   # a file must show something
            next(iter(sheets.values())).hidden = False
        self.disk_sheets = list(sheets)
        self.engine.set_book(sheets)
        self.base, self.history, self.hpos = {}, [], 0
        self.table.sheet = None
        self.sheet = self.visible_sheets()[0]
        self.tabs.rebuild()
        self.show(self.sheet)
        self.refresh()

    def new_document(self):
        self.load_book(None, {"Sheet1": io.Sheet([])})

    def new_window(self):
        """New blank workbook in its own window (a separate process)."""
        if getattr(sys, "frozen", False):          # the built exe
            cmd = [sys.executable]
        else:
            pyw = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
            cmd = [pyw if os.path.exists(pyw) else sys.executable, os.path.abspath(__file__)]
        subprocess.Popen(cmd, close_fds=True)
        return "break"

    def confirm_discard(self):
        """True if it's OK to drop the current data."""
        self.table.commit()
        if not self.dirty():
            return True
        ans = messagebox.askyesnocancel("Unsaved changes", "Save changes before closing this file?")
        if ans is None:
            return False
        return self.save() if ans else True

    def open_dialog(self):
        if not self.confirm_discard():
            return
        path = filedialog.askopenfilename(filetypes=OPEN_TYPES)
        if path:
            self.open_file(path)

    def open_recent(self, path):
        if not os.path.exists(path):
            messagebox.showwarning("Open recent", f"File not found:\n{path}")
            recent.remove(path)
            return
        if self.confirm_discard():
            self.open_file(path)

    def open_file(self, path):
        self.config(cursor="watch")
        self.update_idletasks()
        try:
            sheets = io.load(path)
        except ImportError as e:
            messagebox.showerror("Missing library", f"{e}\n\nRun: pip install -r requirements.txt")
            return False
        except Exception as e:
            messagebox.showerror("Error", f"Could not open file:\n{e}")
            return False
        finally:
            self.config(cursor="")
        if not sheets:
            sheets = {"Sheet1": io.Sheet([])}
        self.load_book(path, sheets)
        recent.add(path)
        return True

    def show(self, name):
        if self.sheet is not None and name != self.sheet:
            self.begin_cross_sheet()
        self.sheet = name
        self.table.load(self.sheets[name])
        if self.formula_home is not None:   # still typing the formula: keep it in the formula bar
            self.value.focus_set()
            self.table.show_refs(self.value.get())

    # ---------- formulas that point at other sheets ----------
    # Switching sheets while typing a formula keeps the formula open in the formula bar, like Excel:
    # clicks on the other sheet insert "Sheet!A1", Enter puts the formula in its original cell.

    def begin_cross_sheet(self):
        g = self.table
        if self.formula_home is not None:
            return
        if g.editor is not None and g.editor.get().startswith("="):
            text, caret = g.editor.get(), g.editor.index("insert")
            _, r, c = g.editor.cell
            g.commit(cancel=True)
        elif self.focus_get() is self.value and self.value.get().startswith("="):
            text, caret = self.value.get(), self.value.index("insert")
            r, c = g.span(*g.cur)[:2]
        else:
            return
        self.formula_home = (self.current, r, c)
        self.value.delete(0, "end")
        self.value.insert(0, text)
        self.value.icursor(caret)

    def home_sheet_name(self):
        """Sheet the formula being typed belongs to (refs without 'Sheet!' point there)."""
        if self.formula_home is not None:
            return self.name_of(self.formula_home[0])
        return self.sheet

    def end_cross_sheet(self, save):
        """Go back to the formula's cell; save=True stores the formula there."""
        home, r, c = self.formula_home
        text = self.value.get().replace(NEWLINE_MARK, "\n")
        self.show(self.name_of(home))
        self.tabs.select(self.sheet)
        self.formula_home = None
        g = self.table
        g.cur = g.anchor = g.ext = (r, c)
        g.point = None
        g.show_refs("")
        if save:
            self.set_cells([(r, c, text)])
            g.cv.focus_set()
            g.move(1, 0)
        else:
            g.cv.focus_set()
            g.redraw()
            self.on_select()

    def truncated(self):
        return [n for n, s in self.sheets.items() if len(s.rows) >= io.MAX_ROWS]

    def save(self):
        """Save to the current file. Returns True on success."""
        self.table.commit()
        if not self.path or os.path.splitext(self.path)[1].lower() not in io.SAVABLE:
            return self.save_as()
        return self.write(self.path)

    def save_as(self):
        self.table.commit()
        stem = os.path.splitext(os.path.basename(self.path))[0] if self.path else "Untitled"
        dst = filedialog.asksaveasfilename(filetypes=SAVE_TYPES, defaultextension=".xlsx",
                                           initialfile=stem + ".xlsx")
        if not dst:
            return False
        return self.write(dst)

    def write(self, dst):
        src_ext = os.path.splitext(self.path)[1].lower() if self.path else ""
        dst_ext = os.path.splitext(dst)[1].lower()
        one_sheet = False
        try:
            if dst_ext in io.XLSX_EXTS and src_ext in io.XLSX_EXTS:
                io.patch_xlsx(self.path, dst, self.changes("v"), self.changes("s"), self.sheets,
                              widths=self.changes("w"))
            elif self.truncated():
                raise ValueError(f"Only the first {io.MAX_ROWS:,} rows were loaded, saving would "
                                 f"lose data. Save as .xlsx from an .xlsx file instead.")
            elif dst_ext in io.XLSX_EXTS:
                io.write_xlsx(dst, self.sheets)
            elif dst_ext in io.CSV_EXTS:
                one_sheet = len(self.sheets) > 1
                if one_sheet and not messagebox.askokcancel(
                        "Only one sheet", f"CSV files hold one sheet. Save only “{self.sheet}”?"):
                    return False
                if self.changes("s") and not messagebox.askokcancel(
                        "Formatting will be lost", "CSV files can't store formatting.\n\n"
                        "Save values only? (Use Save As → .xlsx to keep formatting.)"):
                    return False
                same = self.path is not None and os.path.abspath(dst) == os.path.abspath(self.path)
                g = self.table   # csv stores results, not formulas
                rows = [[g.value(r, c) for c in range(len(row))] for r, row in enumerate(g.sheet.rows)]
                io.write_csv(dst, rows, like=self.path if same else None)
            else:
                raise ValueError(f"Can't save as '{dst_ext}'. Use .xlsx or .csv.")
        except Exception as e:
            messagebox.showerror("Save failed", str(e))
            return False
        recent.add(dst)
        if one_sheet:   # csv kept only the current sheet: show what was saved
            self.open_file(dst)
            return True
        self.path = dst
        self.base = {k: self.read(*k) for k in self.base}
        for group in self.history:      # undo snapshots: re-point their maps at the new file
            for kind, _, _, _, old, new in group:
                if kind in ("x", "b"):
                    structure.rebase(old if kind == "x" else old[1], self.sheets)
                    structure.rebase(new if kind == "x" else new[1], self.sheets)
        live = {id(s) for s in self.sheets.values()}
        for group in self.history:      # deleted sheets only kept for undo are no longer in the file
            for kind, _, _, _, old, new in group:
                if kind == "b":
                    for _, s in old[0] + new[0]:
                        if id(s) not in live:
                            s.disk_name = None
                            s.reset_structure()
        for name, s in self.sheets.items():   # the file now matches the model
            s.reset_structure()
            s.disk_name, s.disk_hidden = name, s.hidden
        self.disk_sheets = list(self.sheets)
        self.refresh()
        self.status["text"] = f"Saved {os.path.basename(dst)}"
        return True

    def on_close(self):
        if self.confirm_discard():
            self.destroy()


if __name__ == "__main__":
    if sys.platform == "win32":  # crisp text on scaled displays
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    app = App()
    if len(sys.argv) > 1:
        app.after(50, app.open_file, sys.argv[1])
    app.mainloop()
