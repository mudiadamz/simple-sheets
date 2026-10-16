"""Simple Sheets - a tiny spreadsheet viewer/editor for .xlsx / .xls / .ods / .csv files."""
import datetime
import json
import os
import subprocess
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

import app_menu
import autocomplete
import charts
import comments
import data_tools
import diskwatch
import dropfiles
import externals
import extras
import fill as fill_series
import find_replace
import formula
import images
import loader
import menus
import recent
import recovery
import rules
import settings
import sheet_io as io
import sheet_props
import structure
import tables
import theme
import workbook_tools
from format_bar import FormatBar, Tooltip
from formatting import DEFAULT_SIZE, parse_number
from formula import Engine
from grid import ALL, Grid, col_name
from sheet_tabs import SheetTabs

OPEN_TYPES = [("Spreadsheets", "*.xlsx *.xlsm *.xlsb *.xls *.ods *.csv *.tsv *.xltx *.xltm"), ("All files", "*.*")]
DROP_EXTS = {".xlsx", ".xlsm", ".xlsb", ".xls", ".ods", ".csv", ".tsv", ".xltx", ".xltm"}
SAVE_TYPES = [("Excel workbook", "*.xlsx"), ("CSV (current sheet)", "*.csv"), ("Excel template", "*.xltx"),
              ("OpenDocument spreadsheet", "*.ods"), ("Excel binary workbook", "*.xlsb")]
LOOK_ONLY = {"s", "w", "rh", "cm", "th", "ln", "fz", "sp", "cf", "dv", "ch", "or", "oc", "m"}   # never change results
STATS_LIMIT = 200_000  # cells scanned for the Sum/Count status
ZOOMS = [0.5, 0.67, 0.75, 0.8, 0.9, 1.0, 1.1, 1.25, 1.5, 1.75, 2.0, 2.5, 3.0]
BAD_NAME_CHARS = set('[]:*?/\\')
NEWLINE_MARK = "↵"   # line breaks as shown in the one-line formula bar


class NameBox(ttk.Entry):
    """The box left of the formula bar. Shows the selection (`box["text"] = "A1"` like a Label)
    unless you're in the middle of typing in it."""

    def __init__(self, master, **kw):
        super().__init__(master, **kw)
        self.typing = False
        self.bind("<Key>", self._key, add="+")
        self.bind("<FocusOut>", lambda e: setattr(self, "typing", False), add="+")

    def _key(self, e):
        self.typing = e.keysym not in ("Return", "KP_Enter", "Escape", "Tab")

    def __setitem__(self, key, value):
        if key != "text":
            return super().__setitem__(key, value)
        if not self.typing:
            self.delete(0, "end")
            self.insert(0, value)

    def __getitem__(self, key):
        return self.get() if key == "text" else super().__getitem__(key)


def resource(*parts):
    """Path of a bundled file, both when run from source and from the built exe."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, *parts)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        import settings_ui
        settings_ui.apply_startup()
        self.title("Simple Sheets")
        try:
            self.iconbitmap(default=resource("assets", "simple_sheets.ico"))
        except tk.TclError:
            pass
        self.scale = self.winfo_fpixels("1i") / 96   # 1.0 at 100% display scaling
        self.scale *= settings_ui.ui_factor(self)      # Settings → Interface size (100% = as Windows has it)
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
        self.names, self.disk_names = {}, {}   # defined names {name: "Sheet1!$A$1:$B$5"}, and as on disk

        theme.base_styles(ttk.Style(self), self.scale)
        self.dark_var = tk.BooleanVar(self, value=bool(settings.get("dark", False)))
        self.contrast_var = tk.BooleanVar(self, value=bool(settings.get("contrast", False)))
        self.grid_var = tk.BooleanVar(self, value=True)        # View menu check marks
        self.headings_var = tk.BooleanVar(self, value=True)
        self.formulas_var = tk.BooleanVar(self, value=False)
        self.full_var = tk.BooleanVar(self, value=False)
        self.page_breaks_var = tk.BooleanVar(self, value=False)
        self.page_breaks = False   # View → Page breaks: dashed lines where printed pages end
        self.show_formulas = False
        self.book_props, self.disk_book_props = {}, {}

        # one toolbar row: ☰ menu, undo/redo, formatting
        top = ttk.Frame(self, padding=(4, 4, 0, 0))
        top.pack(fill="x")
        app_menu.build(self, top)
        app_menu.ai_button(self, top)          # right end of the row (packed before the format bar takes the rest)
        self.fmtbar = FormatBar(top, self, self.scale)
        self.fmtbar.pack(side="left", fill="x", expand=True)

        # formula bar: cell ref | value (editable) | format of the active cell
        fbar = ttk.Frame(self, padding=(4, 0, 4, 4))
        fbar.pack(fill="x")
        self.ref = NameBox(fbar, width=14, justify="center")   # shows the selection; type to go there
        self.ref.pack(side="left", ipady=1)
        self.ref.bind("<Return>", lambda e: (setattr(self.ref, "typing", False), self.goto_ref(self.ref.get()))[-1])
        self.ref.bind("<Escape>", lambda e: (setattr(self.ref, "typing", False), self.on_select(),
                                             self.table.cv.focus_set()))
        self.ref.bind("<FocusIn>", lambda e: self.ref.select_range(0, "end"), add="+")
        Tooltip(self.ref, "Name box: type a cell or range (B12, A1:C5, Sheet2!B3) and press Enter (Ctrl+G)")
        from cell_editor import FormulaBar
        self.value = FormulaBar(fbar, self.table_font_for_bar())
        # ▾ / ▴: show long text on several lines, or keep the bar to one line (like Excel)
        self.fbar_toggle = ttk.Button(fbar, text="▴" if self.value.expanded else "▾", width=2, style="Toolbutton",
                                      command=self.toggle_formula_bar)
        self.fbar_toggle.pack(side="right", anchor="n")
        Tooltip(self.fbar_toggle, "Expand / collapse the formula bar")
        self.value.pack(side="left", fill="x", expand=True, padx=4)
        self.value.bind("<Return>", self.value_commit)
        self.value.bind("<KeyRelease>", lambda e: self.table.show_refs(self.value.get()), add="+")
        self.value.bind("<FocusIn>", lambda e: self.table.show_refs(self.value.get()), add="+")
        self.value.bind("<FocusOut>", lambda e: self.table.show_refs(""), add="+")
        self.value.bind("<Escape>", self.value_cancel)
        autocomplete.attach(self.value, self)

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
        self.status.bind("<Button-1>", lambda e: self.load_all_rows() if "load all" in self.status["text"] else None)

        self.engine = Engine()
        self.rules = rules.Rules(self)   # data validation + conditional formatting from the file
        self.table = Grid(self, self, self.scale)
        self.table.pack(fill="both", expand=True)
        import book_tabs
        book_tabs.setup(self)          # several workbooks in this window (tabs), see book_tabs.py
        for k, key in (("b", "bold"), ("i", "italic"), ("u", "underline"), ("Key-5", "strike")):
            self.table.cv.bind(f"<Control-{k}>", lambda e, key=key: self.fmtbar.shortcut(key))
        self.table.cv.bind("<Control-d>", lambda e: self.fill_copy("down"))
        self.table.cv.bind("<Control-r>", lambda e: self.fill_copy("right"))

        self.bind("<Control-n>", lambda e: self.new_window())
        self.bind("<Control-f>", lambda e: find_replace.open_dialog(self))
        self.bind("<Control-k>", lambda e: (self.edit_link(), "break")[-1])
        self.bind("<Control-g>", lambda e: (self.ref.focus_set(), self.ref.select_range(0, "end"), "break")[-1])
        self.bind("<Control-h>", lambda e: find_replace.open_dialog(self, replace=True))
        self.bind("<Control-o>", lambda e: self.open_dialog())
        self.bind("<Control-s>", lambda e: self.save())
        self.bind("<Control-S>", lambda e: self.save_as())  # Ctrl+Shift+S
        self.bind("<Control-p>", lambda e: (self.print_sheet(), "break")[-1])
        self.bind("<F11>", lambda e: workbook_tools.toggle_full_screen(self))
        self.bind("<Shift-F3>", lambda e: (__import__("function_help").InsertFunction(self), "break")[-1])
        # keyboard-only use: F10 opens the ☰ menu, Ctrl+PgUp / PgDn change sheet, Shift+F10 = right-click menu
        self.bind("<F10>", lambda e: self.post_app_menu())
        self.bind("<Control-Prior>", lambda e: self.next_sheet(-1))
        self.bind("<Control-Next>", lambda e: self.next_sheet(1))
        self.bind("<Shift-F10>", lambda e: self.keyboard_context_menu())
        try:
            self.bind("<App>", lambda e: self.keyboard_context_menu())   # the Menu key (Windows)
        except tk.TclError:
            pass
        self.bind("<Control-e>", lambda e: (self.data_tool("flash_fill"), "break")[-1])
        self.bind("<Control-t>", lambda e: (__import__("book_tabs").new_tab(self), "break")[-1])
        self.bind("<Alt-Shift-Right>", lambda e: (self.group(True), "break")[-1])
        self.bind("<Alt-Shift-Left>", lambda e: (self.group(False), "break")[-1])
        self.bind("<Control-grave>", lambda e: (workbook_tools.toggle_formulas(self), "break")[-1])
        self.bind("<F9>", lambda e: workbook_tools.recalculate(self))
        self.bind("<F1>", lambda e: (self.show_help("shortcuts"), "break")[-1])
        self.bind("<F7>", lambda e: (self.spell_check(), "break")[-1])
        self.bind("<Control-z>", lambda e: self.undo())
        self.bind("<Control-y>", lambda e: self.redo())
        self.bind("<Control-Z>", lambda e: self.redo())  # Ctrl+Shift+Z
        for k in ("plus", "equal", "KP_Add"):
            self.bind(f"<Control-{k}>", lambda e: self.zoom_step(1))
        for k in ("minus", "KP_Subtract"):
            self.bind(f"<Control-{k}>", lambda e: self.zoom_step(-1))
        self.bind("<Control-0>", lambda e: self.zoom_to(1.0))
        self.bind("<Control-greater>", lambda e: (self.fmtbar.step_size(1), "break")[-1])   # Ctrl+Shift+> (Excel)
        self.bind("<Control-less>", lambda e: (self.fmtbar.step_size(-1), "break")[-1])     # Ctrl+Shift+<
        self.bind("<Control-KP_0>", lambda e: self.zoom_to(1.0))
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        # drag files from Explorer: hooked up when the window is shown (its handle can change before that,
        # and asking for it early would also slow down opening the window)
        self.drop_enabled, self._drop_frame = False, None
        self.bind("<Map>", self._hook_drop, add="+")
        self.disk_stamp, self.locked, self.disk_asked = None, False, None
        self.recovery_id, self.autosaved, self.recovered = recovery.new_id(), None, False
        self.after(recovery.INTERVAL_MS, self.autosave_tick)
        self.bind("<Activate>", lambda e: self.after_idle(self.check_disk), add="+")   # back from another app
        self.new_document()
        self.set_theme(self.dark_var.get())
        import i18n  # other languages: dialogs and messages too, not only menus
        i18n.install(self)
        import keyboard_nav  # Ctrl+Alt+5 objects, F6 panes, sheet tabs by keyboard
        keyboard_nav.install(self)
        self.bind("<Control-P>", lambda e: (__import__("palette").open_palette(self), "break")[-1])  # Ctrl+Shift+P
        self.bind("<Control-Shift-A>", lambda e: (__import__("ai_assist").open_assistant(self), "break")[-1])
        import platform_ui  # macOS: ⌘ works wherever Ctrl does
        platform_ui.mirror_command_keys(self)
        platform_ui.mirror_command_keys(self.table.cv)

    def toggle_formula_bar(self):
        on = self.value.set_expanded(not self.value.expanded)
        self.fbar_toggle.configure(text="▴" if on else "▾")
        return on

    def table_font_for_bar(self):
        from tkinter import font as tkfont
        return tkfont.nametofont("TkDefaultFont", root=self)

    def _hook_drop(self, e=None):
        if e is not None and e.widget is not self:
            return
        frame = self.wm_frame()
        if frame != self._drop_frame:
            self._drop_frame = frame
            self.drop_enabled = dropfiles.enable(self, self.open_dropped)

    def report_callback_exception(self, exc_type, exc, tb):
        """An unexpected error in a click / key handler: show it (Copy details) and log it, keep running."""
        import crash
        crash.report(self, exc_type, exc, tb)

    def toggle_dark(self):
        self.set_theme(self.dark_var.get())
        settings.set("dark", self.dark_var.get())

    def toggle_contrast(self):
        self.set_theme(self.dark_var.get())
        settings.set("contrast", self.contrast_var.get())

    def set_theme(self, dark):
        """Light / dark (or high contrast) for the whole window."""
        self.dark_var.set(dark)
        theme.apply(self, dark, self.contrast_var.get())
        self.fmtbar.refresh_icons()
        self.__dict__.pop("menu_icons", None)   # right-click menu icons get redrawn in the new ink
        pane = getattr(self, "ai_pane", None)
        if pane is not None and pane.win.winfo_exists():
            pane.recolor()                       # the AI chat keeps its own colours
        app_menu.refresh_ai_button(self)
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
        if kind == "sp":                   # "sp" = sheet setting `c` (sheet_props.py), None = default
            return sheet.props.get(c)
        if kind == "wp":                   # "wp" = workbook property `c` (title, calc mode…), sheet is None
            return self.book_props.get(c)
        if kind == "cf":                   # "cf" / "dv" = the sheet's conditional formats / validation rules
            return list(sheet.cond)
        if kind == "dv":
            return list(sheet.validations)
        if kind == "ch":                   # "ch" = the sheet's charts
            return list(sheet.charts)
        if kind == "th":                   # "th" = threaded comment of cell (r, c): messages, or None
            return sheet.threads.get((r, c))
        if kind == "im":                   # "im" = the sheet's pictures
            return list(sheet.extra.get("images") or [])
        if kind == "rt":                   # "rt" = rich text runs of cell (r, c) (None = one font)
            return sheet.rich.get((r, c))
        if kind == "tb":                   # "tb" = the sheet's Excel tables (ranges grow as you type)
            return [dict(t) for t in sheet.extra.get("tables", [])]
        if kind == "or":                   # "or" / "oc" = outline (group) level of row r / column c
            return sheet.outline_rows.get(r, 0)
        if kind == "oc":
            return sheet.outline_cols.get(c, 0)
        if kind == "m":                   # "m" = the sheet's merged ranges (whole list)
            return list(sheet.merges)
        if kind == "w":                    # "w" = width of column c (Excel characters; None = default)
            return sheet.widths.get(c)
        if kind == "fz":                   # "fz" = frozen (rows, cols)
            return tuple(sheet.freeze)
        if kind == "flt":                  # "flt" = the sheet's filter (dict or None)
            return sheet.filter
        if kind == "cm":                   # "cm" = comment of cell (r, c): [text, author] or None
            return sheet.comments.get((r, c))
        if kind == "rh":                   # "rh" = height of row r in points (None = automatic)
            return sheet.heights.get(r)
        if kind == "ln":                   # "ln" = hyperlink address of cell (r, c)
            return sheet.links.get((r, c))
        if kind == "nm":                   # "nm" = the workbook's defined names (sheet is None)
            return dict(self.names)
        if kind == "hr":
            return r in sheet.hidden_rows
        if kind == "hc":
            return c in sheet.hidden_cols
        return self.get(sheet, r, c) if kind == "v" else sheet.styles.get((r, c))

    def write_cell(self, kind, sheet, r, c, val):
        if kind in ("sp", "wp"):
            target = sheet.props if kind == "sp" else self.book_props
            if val is None:
                target.pop(c, None)
            else:
                target[c] = val
            if kind == "wp" and c in ("iterate", "iterate_count", "iterate_delta"):
                self.engine.set_iteration(self.book_props)
        elif kind == "cf":
            sheet.cond = list(val)
        elif kind == "dv":
            sheet.validations = list(val)
        elif kind == "ch":
            sheet.charts = list(val)
        elif kind == "th":
            if val:
                sheet.threads[(r, c)] = [list(m) for m in val]
            else:
                sheet.threads.pop((r, c), None)
        elif kind == "im":
            sheet.extra["images"] = list(val)
        elif kind == "rt":
            if val:
                sheet.rich[(r, c)] = [list(x) for x in val]
            else:
                sheet.rich.pop((r, c), None)
        elif kind == "tb":
            sheet.extra["tables"] = [dict(t) for t in val]
            self.engine.set_tables([dict(t, sheet=n) for n, s in self.sheets.items() for t in s.extra.get("tables", [])])
        elif kind in ("or", "oc"):
            target, i = (sheet.outline_rows, r) if kind == "or" else (sheet.outline_cols, c)
            if val:
                target[i] = val
            else:
                target.pop(i, None)
        elif kind == "nm":
            self.names = dict(val)
            self.engine.set_names(self.names)
        elif kind == "ln":
            if val:
                sheet.links[(r, c)] = val
            else:
                sheet.links.pop((r, c), None)
        elif kind == "rh":
            if val is None:
                sheet.heights.pop(r, None)
            else:
                sheet.heights[r] = val
        elif kind == "cm":
            if val:
                sheet.comments[(r, c)] = list(val)
            else:
                sheet.comments.pop((r, c), None)
        elif kind == "fz":
            sheet.freeze = tuple(val)
        elif kind == "flt":
            sheet.filter = val
            data_tools.recompute(self, self.name_of(sheet))
        elif kind == "m":
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
            self.engine.note(sheet, r, c, val)
        elif val:
            sheet.styles[(r, c)] = val
        else:
            sheet.styles.pop((r, c), None)

    def apply(self, kind, changes):
        """Apply [(r, c, new), ...] on the current sheet as one undo step."""
        self.apply_mixed([(kind, r, c, val) for r, c, val in changes])

    def apply_mixed(self, changes):
        """Apply [(kind, r, c, new), ...] (values and styles) on the current sheet as one undo step."""
        self.apply_cells([(kind, self.current, r, c, val) for kind, r, c, val in changes])

    def apply_cells(self, changes):
        """Apply [(kind, Sheet, r, c, new), ...] - possibly on several sheets - as one undo step."""
        if any(s is not None and s.props.get("protect") and kind not in ("sp", "wp")
               and not (kind == "v" and not workbook_tools.is_locked(self, s, r, c))
               for kind, s, r, c, _ in changes):
            self.protected_warning()
            return
        group = []
        for kind, s, r, c, val in changes:
            old = self.read(kind, s, r, c)
            if val != old:
                self.base.setdefault((kind, s, r, c), old)
                group.append((kind, s, r, c, old, val))
                self.write_cell(kind, s, r, c, val)
        if group:
            self.push(group)

    def protected_warning(self):
        messagebox.showinfo("Protected sheet", "This sheet is protected: only unlocked cells can be changed.\n\n"
                            "Use ☰ → Sheet → Unprotect sheet to change it.", parent=self)

    def sheet_protected(self, name, action):
        """A protected sheet can't be deleted / renamed / hidden (Excel would allow it; it surprises people)."""
        if name in self.sheets and self.sheets[name].props.get("protect"):
            messagebox.showinfo("Protected sheet", f"“{name}” is protected, so it can't be {action}.\n\n"
                                "Unprotect it first (right-click its tab → Unprotect sheet).", parent=self)
            return True
        return False

    def push(self, group):
        del self.history[self.hpos:]
        self.history.append(group)
        self.hpos += 1
        self.after_change(group)

    def restructure(self, axis, at, n):
        """Insert (n > 0) or delete (n < 0) rows/columns at `at` on the current sheet."""
        self.table.commit()
        if self.current.props.get("protect"):
            self.protected_warning()
            return
        if self.sheet in self.truncated():
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
        group = [("x", self.current, r, c, before, structure.snapshot(self.sheets))]
        names = {k: formula.adjust(t, self.sheet, self.sheet, axis, at, n) for k, t in self.names.items()}
        if names != self.names:   # defined names follow too (same undo step)
            group.append(("nm", None, 0, 0, dict(self.names), names))
            self.write_cell("nm", None, 0, 0, names)
        self.push(group)

    def move_lines(self, axis, a, n, t):
        """Move rows / columns a .. a+n-1 so they sit before line t (dragging selected headers). One undo
        step; the moved lines stay selected. Returns the new first index, or None if nothing moved."""
        self.table.commit()
        if a <= t <= a + n:
            return None
        if self.current.props.get("protect"):
            self.protected_warning()
            return None
        if self.sheet in self.truncated():
            messagebox.showwarning("Too big", f"Only the first {io.MAX_ROWS:,} rows are loaded, so rows "
                                              "and columns can't be moved in this sheet.")
            return None
        before = structure.snapshot(self.sheets)
        try:
            changed = structure.move(self.sheets, self.sheet, axis, a, n, t)
        except ValueError:
            structure.restore(self.sheets, before)
            messagebox.showwarning("Move", "Part of a merged block would move away from the rest. "
                                           "Select the whole merged rows / columns, or unmerge first.", parent=self)
            return None
        for home, r, c in changed:
            if home != self.sheet:   # formulas on other sheets: saved as normal cell edits
                self.base.setdefault(("v", self.sheets[home], r, c), before[home]["rows"][r][c])
        first = t if t < a else t - n
        g = self.table
        cur = (0, first) if axis == "col" else (first, 0)
        group = [("x", self.current, cur[0], cur[1], before, structure.snapshot(self.sheets))]
        perm = structure.mover(a, n, t)
        names = {k: formula.permute(v, self.sheet, self.sheet, axis, perm) for k, v in self.names.items()}
        if names != self.names:
            group.append(("nm", None, 0, 0, dict(self.names), names))
            self.write_cell("nm", None, 0, 0, names)
        self.push(group)
        g.select_line(axis, first, first + n - 1)
        return first

    def shift_cells(self, direction):
        """Insert cells (shift right/down) or delete cells (shift left/up) for the selection."""
        g = self.table
        g.commit()
        r1, c1, r2, c2 = g.sel_cells()
        self.apply_mixed(structure.shift_cells(g.sheet, (r1, c1, r2, c2), direction))
        g.select(r1, c1, r2, c2)

    def set_cells(self, changes):
        import tables
        extra = tables.grow(self, changes)   # typing under an Excel table extends it (same undo step)
        if extra:
            self.apply_mixed([("v", r, c, v) for r, c, v in changes] + [extra])
        else:
            self.apply("v", changes)

    def add_reply(self, r, c, text, start=False):
        """A threaded comment: start one (start=True) or reply to the cell's conversation."""
        import threads
        old = self.current.threads.get((r, c)) or []
        if not old and not start:
            return False
        msg = [comments.author(), text, threads.now(), threads.new_id()]
        thread = [list(m) for m in old] + [msg]
        self.apply_mixed([("th", r, c, thread), ("cm", r, c, [threads.fallback(thread), "tc=" + thread[0][3]])])
        return True

    def delete_thread(self, r, c):
        self.apply_mixed([("th", r, c, None), ("cm", r, c, None)])

    def set_comment(self, r, c, text):
        """Add / change a cell's comment; empty text deletes it."""
        old = self.current.comments.get((r, c))
        new = [text, old[1] if old else comments.author()] if text else None
        self.apply_mixed([("cm", r, c, new)])

    def set_names(self, names):
        """Replace the defined names (one undo step)."""
        self.apply_cells([("nm", None, 0, 0, dict(names))])

    def manage_names(self):
        import name_manager
        self.table.commit()
        return name_manager.NameManager(self)

    def set_link(self, r, c, text, url):
        """Text + link for a cell (url None removes the link). A new link gets the usual blue underline."""
        s = self.current
        changes = [("v", r, c, text), ("ln", r, c, url)]
        st = dict(s.styles.get((r, c)) or {})
        if url and (r, c) not in s.links:
            changes.append(("s", r, c, {**st, "fg": "#0563c1", "underline": True}))
        elif not url and (r, c) in s.links and st.get("fg") == "#0563c1":
            changes.append(("s", r, c, {k: v for k, v in st.items() if k not in ("fg", "underline")} or None))
        self.apply_mixed(changes)

    def edit_link(self):
        g = self.table
        g.commit()
        return comments.LinkDialog(self, *g.span(*g.cur)[:2])

    def edit_comment(self):
        g = self.table
        g.commit()
        return comments.CommentDialog(self, *g.span(*g.cur)[:2])

    def open_filter(self, c, x_root, y_root):
        """▼ button of a filter column."""
        old = getattr(self, "filter_popup", None)
        if old is not None and old.winfo_exists():
            old.destroy()
        self.filter_popup = data_tools.FilterPopup(self, c, x_root, y_root)

    def set_freeze(self, rows, cols):
        """Freeze the top `rows` rows and left `cols` columns (0, 0 = unfreeze)."""
        self.table.commit()
        self.apply_mixed([("fz", 0, 0, (max(rows, 0), max(cols, 0)))])

    def set_height(self, r, points):
        """Row height in points (None = automatic) as an undoable, saved change."""
        self.apply_mixed([("rh", r, 0, points)])

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
            self.format_rect(self.table.cells_rects(), lambda st, r, c: fn(st))

    def format_rect(self, rect, fn):
        """Restyle a block (or a list of blocks, one undo step): fn(style dict, r, c) -> new style dict."""
        g = self.table
        rects = rect if isinstance(rect, list) else [rect]
        if sum((r2 - r1 + 1) * (c2 - c1 + 1) for r1, c1, r2, c2 in rects) > 500_000:
            messagebox.showwarning("Too many cells", "Select a smaller range to format.")
            return
        changes, seen = [], set()
        for r1, c1, r2, c2 in rects:
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    if g.span(r, c)[:2] != (r, c) or (r, c) in seen:
                        continue   # merged-away cell: the block's top-left holds the format
                    seen.add((r, c))
                    st = fn(dict(g.sheet.styles.get((r, c)) or {}), r, c)
                    st = {k: v for k, v in st.items() if v and not (k == "size" and v == DEFAULT_SIZE)}
                    changes.append((r, c, st or None))
        self.apply("s", changes)

    def after_change(self, group):
        kinds = {g[0] for g in group}
        if self.book_props.get("calc") == "manual" and kinds <= {"v", "s"}:
            for kind, s, r, c, _, _new in group:   # manual mode: only the typed cells get a new value (F9 = all)
                if kind == "v":
                    name = self.name_of(s)
                    self.engine.memo.pop((name, r, c), None)
        elif kinds <= {"v", "rt"}:     # typed values: recalculate only what depends on those cells
            names = {id(s): n for n, s in self.sheets.items()}
            self.engine.invalidate({(names[id(s)], r, c) for kind, s, r, c, _, _ in group
                                    if kind == "v" and id(s) in names})
        elif kinds <= LOOK_ONLY:       # formatting / layout: no result changes
            pass
        else:
            self.engine.clear()   # recalculate formulas on next draw
        if "b" in kinds:
            self.after_book_change()
        elif "x" in kinds:
            self.engine.rescan()
            r, c = group[0][2], group[0][3]
            self.table.reload()
            self.table.goto(r, c)
        elif "s" in kinds:
            self.table.style_changed({g[2] for g in group})
        if kinds & {"hr", "hc", "m", "w", "fz", "flt", "rh", "or", "oc", "sp"}:
            self.table.reload()
            self.table.fix_cursor()
        if "sp" in kinds:
            self.tabs.rebuild()          # tab colours
        self.table.data_changed()
        if getattr(self, "split_grid", None) is not None:
            self.split_grid.sync()
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
        return getattr(self, "recovered", False) or getattr(self, "password_changed", False) or self.book_changed() or \
            getattr(self, "book_props", {}) != getattr(self, "disk_book_props", {}) or \
            any(s.structure_changed() for s in self.sheets.values()) or \
            self.names != self.disk_names or \
            any(self.read(*k) != v for k, v in self.base.items() if k[1] is not None and id(k[1]) in live)

    def refresh(self):
        self.undo_btn["state"] = "normal" if self.hpos > 0 else "disabled"
        self.redo_btn["state"] = "normal" if self.hpos < len(self.history) else "disabled"
        name = os.path.basename(self.path) if self.path else "Untitled"
        dirty = self.dirty()
        note = "  (in use by another program - save as a copy)" if getattr(self, "locked", False) else ""
        self.title(f"Simple Sheets - {'*' if dirty else ''}{name}{note}")
        self.save_btn["state"] = "normal" if dirty else "disabled"
        self.on_select()

    # ---------- sheets: add / rename / delete ----------

    def book_state(self, select=None):
        """(sheet order, contents snapshot, sheet to show, defined names) - one side of a sheet
        add/rename/delete."""
        return list(self.sheets.items()), structure.snapshot(self.sheets), select or self.sheet, dict(self.names)

    def restore_book(self, state):
        order, snap, self.select_after, names = state
        self.sheets.clear()
        self.sheets.update(order)
        structure.restore(self.sheets, snap)
        self.names = dict(names)
        self.engine.set_names(self.names)

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
        if old != new and self.sheet_protected(old, "renamed"):
            return False
        err = self.check_name(new, old)
        if err:
            messagebox.showwarning("Rename sheet", err)
            return False

        def mutate():
            items = [(new if n == old else n, s) for n, s in self.sheets.items()]
            self.sheets.clear()
            self.sheets.update(items)
            self.fix_formulas(lambda t: formula.rename_sheet(t, old, new))
            self.names = {k: formula.rename_sheet(t, old, new) for k, t in self.names.items()}
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
        if hide and self.sheet_protected(name, "hidden"):
            return False
        if hide and self.visible_sheets() == [name]:
            messagebox.showwarning("Hide sheet", "At least one sheet must stay visible.")
            return False
        select = self.neighbour(name) if hide else name

        def mutate():
            self.sheets[name].hidden = hide
        self.book_change(mutate, select)
        return True

    def delete_sheet(self, name):
        if self.sheet_protected(name, "deleted"):
            return False
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
            self.names = {k: t for k, t in self.names.items() if formula.rename_sheet(t, name, None) == t}
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
        self.fmtbar.sync(g.sheet.styles.get((r, c)))
        self.fmtbar.merged.set((r, c) in g.merge_at)
        # Excel-style stats for multi-cell selections
        if g.multi():
            count, nums, total = 0, 0, 0.0
            blocks = [(a, b, min(c_, len(g.sheet.rows) - 1), min(d, g.ncols - 1)) for a, b, c_, d in g.rects()]
            if sum(max(c_ - a + 1, 0) * max(d - b + 1, 0) for a, b, c_, d in blocks) <= STATS_LIMIT:
                seen = set()
                for dr1, dc1, dr2, dc2 in blocks:
                    for rr in range(dr1, dr2 + 1):
                        for cc in range(dc1, dc2 + 1):
                            if (rr, cc) in seen:
                                continue
                            seen.add((rr, cc))
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
            limit = getattr(g.sheet, "loaded_limit", None)
            note = f" (first {limit:,} only - click to load all)" if limit and rows >= limit else ""
            self.status["text"] = f"{rows:,} rows{note}"
            self.status.configure(cursor="hand2" if note else "")

    def value_commit(self, _=None):
        g = self.table
        g.point = None
        if self.formula_home is not None:
            self.end_cross_sheet(save=True)
            return "break"
        if g.sheet:
            r, c = g.span(*g.cur)[:2]
            val = self.value.get().replace(NEWLINE_MARK, "\n")
            if val != g.text(r, c) and not self.allowed(r, c, val):
                return "break"                    # rejected by data validation: keep typing
            self.set_cells([(r, c, val)])
            g.cv.focus_set()
            g.move(1, 0)
        return "break"

    def allowed(self, r, c, val):
        """Data validation on typed input: 'stop' rejects, 'warning' asks, 'information' just tells."""
        problem = self.rules.check(self.table.sheet, r, c, val)
        if problem is None:
            return True
        style, title, msg = problem
        if style == "warning":
            return messagebox.askokcancel(title, msg + "\n\nKeep this value anyway?", icon="warning", parent=self)
        if style == "information":
            messagebox.showinfo(title, msg, parent=self)
            return True
        messagebox.showerror(title, msg, parent=self)
        return False

    def goto_ref(self, text):
        """Name box / Ctrl+G: jump to a cell, select a range, rows or columns, maybe on another sheet."""
        g = self.table
        target = self.resolve_ref(text)
        if target is None:
            messagebox.showwarning("Go to", f"“{text.strip()}” isn't a cell or range.")
            self.on_select()
            return "break"
        sheet, r1, c1, r2, c2 = target
        if sheet != self.sheet:
            self.show(sheet)
            self.tabs.select(sheet)
        r2 = ALL if r2 >= formula.MAXR - 1 else r2
        c2 = ALL if c2 >= formula.MAXC - 1 else c2
        if (r1, c1) == (r2, c2):
            g.goto(r1, c1)
        else:
            g.select(r1, c1, r2, c2)
        g.cv.focus_set()
        return "break"

    def resolve_ref(self, text):
        """'B12', 'a1:c5', 'B:B', 'Sheet2!B3', a defined name -> (sheet, r1, c1, r2, c2) or None."""
        toks = [(k, t) for k, t in formula.lex(text.strip()) if k != "ws"]
        if len(toks) != 1:
            return None
        kind, tok = toks[0]
        if kind == "name" and tok.lower() in {n.lower() for n in getattr(self, "names", {})}:
            return self.resolve_ref(next(v for n, v in self.names.items() if n.lower() == tok.lower()))
        if kind not in ("ref", "cols", "rows"):
            return None
        sheet, r1, c1, r2, c2 = formula.parse_ref(kind, tok)
        if sheet is not None:
            sheet = next((n for n in self.sheets if n.lower() == sheet.lower()), None)
            if sheet is None or self.sheets[sheet].hidden:
                return None
        return (sheet or self.sheet), r1, c1, r2, c2

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

    def load_book(self, path, sheets, names=None, book_props=None):
        """Show a workbook (sheets as on disk) in this window."""
        self.path, self.sheets = path, sheets
        self.formula_home = None
        self.book_props, self.disk_book_props = dict(book_props or {}), dict(book_props or {})
        self.engine.set_iteration(self.book_props)
        if path and path.lower().endswith(".xlsb"):
            import xlsb
            self.lost_parts = xlsb.unsupported(path)
        else:
            self.lost_parts = extras.lost_parts(path) if path and path.lower().endswith((".xlsx", ".xlsm")) else []
        self.lost_ok = False   # asked once before a save would drop them
        self.names, self.disk_names = dict(names or {}), dict(names or {})
        self.engine.set_names(self.names)
        for name, s in sheets.items():
            s.disk_name, s.disk_hidden = name, s.hidden
            rules.remember(s)
            charts.remember(s)
            tables.remember(s)
            images.remember(s)
        self.engine.set_tables([dict(t, sheet=n) for n, s in sheets.items() for t in s.extra.get("tables", [])])
        self.engine.externals = externals.Links(self, path).value
        self.engine.rescan()
        if all(s.hidden for s in sheets.values()):   # a file must show something
            next(iter(sheets.values())).hidden = False
        self.disk_sheets = list(sheets)
        self.engine.set_book(sheets)
        for name, s in sheets.items():
            if s.filter:   # rows the file's filter hides are filtered, not hidden by hand
                data_tools.recompute(self, name)
                s.hidden_rows -= s.filtered_rows
        self.base, self.history, self.hpos = {}, [], 0
        self.table.sheet = None
        self.sheet = self.visible_sheets()[0]
        self.tabs.rebuild()
        self.show(self.sheet)
        self.disk_stamp = diskwatch.stamp(path) if path else None
        self.locked = diskwatch.is_locked(path)
        self.recovered, self.recovered_from = False, None
        self.template_src = self.encrypted_from = None
        self.plain_src = self.password = None
        self.pivots = []                      # pivot tables made here (pivots.py), written on save
        self.password_changed = False
        self.refresh()

    def new_document(self):
        self.load_book(None, {"Sheet1": io.Sheet([])})

    def new_window(self, path=None, recover=False):
        """New blank workbook (or `path`) in its own window (a separate process)."""
        if getattr(sys, "frozen", False):          # the built exe
            cmd = [sys.executable]
        else:
            pyw = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
            cmd = [pyw if os.path.exists(pyw) else sys.executable, os.path.abspath(__file__)]
        subprocess.Popen(cmd + (["--recover"] if recover else []) + ([path] if path else []), close_fds=True)
        return "break"

    def check_disk(self):
        """Window got focus back: did another program change our file meanwhile? Offer to reload
        (asked once per change). Also notices when Excel opens / closes the file."""
        if not self.path or getattr(self, "checking_disk", False):
            return
        self.checking_disk = True
        try:
            self._check_disk()
        finally:
            self.checking_disk = False

    def _check_disk(self):
        locked = diskwatch.is_locked(self.path)
        if locked != self.locked:
            self.locked = locked
            self.refresh()
        now = diskwatch.stamp(self.path)
        if now is None or self.disk_stamp is None or now == self.disk_stamp or now == self.disk_asked:
            return
        self.disk_asked = now
        lost = "\n\nYour unsaved changes here would be lost." if self.dirty() else ""
        if messagebox.askokcancel("File changed on disk", f"“{os.path.basename(self.path)}” was changed by "
                                  f"another program.\n\nReload it?{lost}", icon="warning"):
            self.table.commit(cancel=True)
            self.open_file(self.path)

    def open_dropped(self, paths):
        """Files dropped on the window: the first opens here if this window is still an empty new one,
        everything else in new windows (nothing unsaved is ever replaced)."""
        paths = [p for p in paths if os.path.isfile(p) and os.path.splitext(p)[1].lower() in DROP_EXTS]
        if not paths:
            messagebox.showwarning("Open", "Drop a spreadsheet file (.xlsx, .xls, .ods, .csv).")
            return
        if self.path is None and not self.dirty():
            self.open_file(paths.pop(0))
        for p in paths:
            self.open_elsewhere(p)

    def confirm_discard(self):
        """True if it's OK to drop the current data."""
        self.table.commit()
        if not self.dirty():
            return True
        ans = messagebox.askyesnocancel("Unsaved changes", "Save changes before closing this file?")
        if ans is None:
            return False
        return self.save() if ans else True

    def tabs_mode(self):
        """Settings → open files in tabs (one window) instead of replacing / new windows."""
        return bool(settings.get("open_in_tabs", False))

    def open_elsewhere(self, path):
        """Another file while this one stays open: a new tab or a new window (Settings)."""
        import book_tabs
        return book_tabs.new_tab(self, path) if self.tabs_mode() else self.new_window(path)

    def open_dialog(self):
        if self.tabs_mode() and (self.path is not None or self.dirty()):
            import book_tabs
            return book_tabs.open_in_tab(self)
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
        if self.tabs_mode() and (self.path is not None or self.dirty()):
            return self.open_elsewhere(path)
        if self.confirm_discard():
            self.open_file(path)

    def unlock(self, path, password=None):
        """A password-protected workbook: ask for the password, decrypt into a private temp file.
        Returns the temp file's path, or None if cancelled."""
        import tempfile

        import office_crypto
        prompt = f"“{os.path.basename(path)}” is protected with a password.\n\nPassword:"
        while True:
            if password is None:
                password = simpledialog.askstring("Password", prompt, show="•", parent=self)
                if password is None:
                    return None
            try:
                data = office_crypto.decrypt(path, password)
                break
            except office_crypto.WrongPassword:
                password = None
                prompt = "The password isn't right. Try again:"
            except Exception as e:
                messagebox.showerror("Open", f"Couldn't decrypt the file:\n{e}", parent=self)
                return None
        fd, tmp = tempfile.mkstemp(suffix=os.path.splitext(path)[1] or ".xlsx", prefix="ss-unlocked-")
        with os.fdopen(fd, "wb") as f:
            f.write(data)
        self.unlocked_files = getattr(self, "unlocked_files", []) + [tmp]
        self.unlocked_password = password
        return tmp

    def open_file(self, path, password=None):
        import office_crypto
        real = path
        if path.lower().endswith((".xlsx", ".xlsm", ".xltx", ".xltm")) and office_crypto.is_encrypted(path):
            path = self.unlock(path, password)
            if path is None:
                return False
        ok = self._open_file(path)
        if ok and real != path:   # opened from a decrypted copy: saving encrypts again with the same password
            self.path, self.plain_src, self.password = real, path, self.unlocked_password
            self.encrypted_from = real
            self.disk_stamp = diskwatch.stamp(real)
            recent.remove(path)
            recent.add(real)
            self.refresh()
            self.status["text"] = "Password-protected file: saved with the same password"
        return ok

    def set_password(self, password=None):
        """Tools → Password: protect the workbook file with a password (empty removes it)."""
        if password is None:
            password = simpledialog.askstring("Password", "Password to open this workbook (empty = none):",
                                              show="•", parent=self)
            if password is None:
                return False
        self.password = password or None
        self.password_changed = True
        self.refresh()
        self.status["text"] = ("Saving will protect the file with a password" if password
                               else "Saving will remove the password")
        return True

    def _open_file(self, path):
        self.table.commit()
        self.config(cursor="watch")
        self.update_idletasks()
        try:
            sheets = loader.load(self, path)       # big files: in the background, with progress
            if sheets is None:
                self.status["text"] = "Opening cancelled"
                return False
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
        for s in sheets.values():
            s.loaded_limit = io.MAX_ROWS   # rows beyond this weren't read (see truncated / load_all_rows)
        if path.lower().endswith(io.XLSX_EXTS):
            try:
                import threads
                for name, found in threads.read(path).items():
                    if name in sheets:
                        sheets[name].threads = found
            except Exception:
                pass
        try:
            names = io.load_names(path)
        except Exception:
            names = {}   # unreadable names never block opening the file
        try:
            book_props = sheet_props.read_book(path) if path.lower().endswith((".xlsx", ".xlsm", ".xltx", ".xltm")) else {}
        except Exception:
            book_props = {}
        self.load_book(path, sheets, names, book_props)
        recent.add(path)
        try:   # OneDrive / Google Drive / Dropbox folder: say so, warn about conflict copies
            import cloud
            if cloud.service(path):
                self.after_idle(lambda: cloud.check(self, path))
        except Exception:
            pass
        return True

    def show(self, name):
        if self.sheet is not None and name != self.sheet:
            self.begin_cross_sheet()
        self.sheet = name
        self.table.load(self.sheets[name])
        if getattr(self, "split_grid", None) is not None:
            self.split_grid.sync()
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
        """Sheets cut at the row limit they were loaded with (Sheet.loaded_limit)."""
        return [n for n, s in self.sheets.items() if getattr(s, "loaded_limit", None) and len(s.rows) >= s.loaded_limit]

    def save(self):
        """Save to the current file. Returns True on success."""
        self.table.commit()
        if not self.path or os.path.splitext(self.path)[1].lower() not in io.SAVABLE:
            return self.save_as()
        return self.write(self.path)

    def save_as(self):
        self.table.commit()
        named = self.path or getattr(self, "recovered_from", None) or getattr(self, "encrypted_from", None)
        stem = os.path.splitext(os.path.basename(named))[0] if named else "Untitled"
        dst = filedialog.asksaveasfilename(filetypes=SAVE_TYPES, defaultextension=".xlsx",
                                           initialfile=stem + ".xlsx")
        if not dst:
            return False
        return self.write(dst)

    def write(self, dst):
        # what the changes are patched into: the decrypted copy of a password-protected file, the file itself,
        # or the template a new workbook came from
        src = getattr(self, "plain_src", None) or self.path or getattr(self, "template_src", None)
        src_ext = os.path.splitext(src)[1].lower() if src else ""
        dst_ext = os.path.splitext(dst)[1].lower()
        one_sheet = False
        if diskwatch.is_locked(dst):
            messagebox.showerror("File in use", f"“{os.path.basename(dst)}” is open in another program "
                                 "(Excel?).\n\nClose it there and save again, or use Save as… to save a copy.")
            return False
        same = self.path is not None and os.path.abspath(dst) == os.path.abspath(self.path)
        now = diskwatch.stamp(dst) if same else None
        if same and now is not None and self.disk_stamp is not None and now != self.disk_stamp and \
                not messagebox.askokcancel("File changed on disk",
                                           f"“{os.path.basename(dst)}” was changed by another program after "
                                           "it was opened here.\n\nSave anyway and replace those changes?",
                                           icon="warning"):
            return False
        lost = getattr(self, "lost_parts", [])
        if lost and not self.lost_ok and (dst_ext in io.XLSX_EXTS and src_ext in io.XLSX_EXTS
                                          or dst_ext == src_ext == ".xlsb"):
            if not messagebox.askokcancel(
                    "Some parts can't be kept", "This file has " + ", ".join(lost) + ", which Simple Sheets "
                    "can't keep: the saved file won't have them.\n\nSave anyway? (Cancel, then Save as… to "
                    "keep the original file as it is.)", icon="warning", parent=self):
                return False
            self.lost_ok = True
        target = dst
        password = getattr(self, "password", None)
        if password and dst_ext in io.XLSX_EXTS:   # written plain to a private temp file, then encrypted
            import tempfile
            fd, target = tempfile.mkstemp(suffix=dst_ext, prefix="ss-plain-")
            os.close(fd)
            self.unlocked_files = getattr(self, "unlocked_files", []) + [target]
        try:
            if dst_ext == ".ods":
                if self.truncated():
                    raise ValueError(f"Only the first {io.MAX_ROWS:,} rows were loaded, saving would lose data.")
                import ods_write
                ods_write.write(dst, self, self.sheets)
            elif dst_ext == ".xlsb":
                if self.truncated():
                    raise ValueError(f"Only the first {io.MAX_ROWS:,} rows were loaded, saving would lose data.")
                import xlsb_write
                n = xlsb_write.write(dst, self, self.sheets)
                if n:
                    messagebox.showwarning("Saved as .xlsb", f"{n} formula{'s' if n != 1 else ''} (other sheets, "
                                           "names or newer functions) were saved as values in the .xlsb file. "
                                           "Save as .xlsx to keep every formula.", parent=self)
            elif dst_ext in io.XLSX_EXTS and src_ext in io.XLSX_EXTS:
                io.patch_xlsx(src, target, self.changes("v"), self.changes("s"), self.sheets,
                              widths=self.changes("w"), freeze=self.changes("fz"),
                              filters=self.changes("flt"), comments=self.changes("cm"),
                              heights=self.changes("rh"), links=self.changes("ln"),
                              names=self.names if self.names != self.disk_names else None,
                              settings=self.changes("sp"),
                              outline={**self.changes("or"), **{("c",) + k: v for k, v in self.changes("oc").items()}},
                              book_props=self.book_props if self.book_props != self.disk_book_props else None,
                              extra=self.save_extras())
            elif self.truncated():
                raise ValueError(f"Only the first {io.MAX_ROWS:,} rows were loaded, saving would "
                                 f"lose data. Save as .xlsx from an .xlsx file instead.")
            elif dst_ext in io.XLSX_EXTS:
                io.write_xlsx(target, self.sheets, self.names, self.book_props)
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
        if dst_ext in io.XLSX_EXTS:
            try:   # formula results into the file, for viewers that don't calculate
                io.write_cached_values(target, self.formula_results())
            except Exception:
                pass   # the file is saved and valid either way; Excel recalculates on open
            try:   # threaded comments (openpyxl drops them): written back from the model
                import threads
                threads.write(target, self.sheets)
            except Exception as e:
                messagebox.showwarning("Save", f"The comment conversations couldn't be saved:\n{e}")
            try:   # pivot tables made here: Excel pivot parts (Excel rebuilds them on open)
                import pivots
                pivots.write(target, self)
            except Exception as e:
                messagebox.showwarning("Save", f"The pivot table couldn't be saved as an Excel pivot:\n{e}")
        if target != dst:
            import office_crypto
            try:
                office_crypto.encrypt_file(target, dst, password)
            except Exception as e:
                messagebox.showerror("Save failed", f"Couldn't protect the file with its password:\n{e}")
                return False
            self.plain_src = target            # the next save patches this plain copy
        elif dst_ext in io.XLSX_EXTS:
            self.plain_src = None
        self.password_changed = False
        recent.add(dst)
        if one_sheet:   # csv kept only the current sheet: show what was saved
            self.open_file(dst)
            return True
        self.path = dst
        self.template_src = None
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
            rules.after_save(s)
            charts.after_save(s)
            tables.remember(s)
            images.after_save(s)
            s.disk_name, s.disk_hidden = name, s.hidden
        self.disk_sheets = list(self.sheets)
        self.disk_names = dict(self.names)
        self.disk_stamp, self.locked = diskwatch.stamp(dst), False
        self.recovered = False
        self.disk_book_props = dict(self.book_props)
        recovery.discard(self.recovery_id)    # saved: the crash-recovery copy isn't needed
        self.refresh()
        self.status["text"] = f"Saved {os.path.basename(dst)}"
        return True

    def formula_results(self):
        """{sheet: {(r, c): calculated value}} for every formula cell (for write_cached_values).
        Fresh calculation, row by row (so long chains build up without deep recursion)."""
        self.engine.clear()
        out = {}
        for name, s in self.sheets.items():
            cells = {}
            for r, row in enumerate(s.rows):
                for c, t in enumerate(row):
                    if len(t) > 1 and t[0] == "=":
                        v = self.engine.cell(name, r, c)
                        if v == formula.NAME and (r, c) in s.cache:   # unknown function: Excel's result
                            v = formula.literal(s.cache[(r, c)])
                        if v is None or v == formula.CIRC:
                            continue
                        cells[(r, c)] = str(v) if isinstance(v, formula.XLError) else \
                            float(v) if isinstance(v, formula.Serial) else v
            if cells:
                out[name] = cells
        return out

    def on_close(self):
        import book_tabs
        if book_tabs.count(self) > 1 and not book_tabs.close_all(self):
            return
        if self.confirm_discard():
            recovery.discard(self.recovery_id)   # closed on purpose: nothing to recover
            for tmp in getattr(self, "unlocked_files", []):   # decrypted copies of password-protected files
                try:
                    os.remove(tmp)
                except OSError:
                    pass
            self.destroy()

    def load_more_rows(self, sheet=None, chunk=None):
        """A long .xlsx sheet scrolled near the end of what's loaded: read the next rows (no reload,
        unsaved changes stay). Returns how many rows were added."""
        sheet = sheet or self.current
        limit = getattr(sheet, "loaded_limit", None)
        src = getattr(self, "plain_src", None) or self.path
        if not limit or len(sheet.rows) < limit or not src or not src.lower().endswith(io.XLSX_EXTS):
            return 0
        if getattr(self, "_loading_more", False):
            return 0
        self._loading_more = True
        chunk = chunk or max(io.MAX_ROWS // 2, 1000)
        try:
            name = sheet.disk_name or self.name_of(sheet)
            self.config(cursor="watch")
            self.update_idletasks()
            rows, styles, cache = io.load_xlsx_rows(src, name, limit, limit + chunk)
        except Exception:
            return 0
        finally:
            self._loading_more = False
            self.config(cursor="")
        first = len(sheet.rows)
        sheet.rows.extend(rows)
        sheet.styles.update(styles)
        sheet.styles0.update(styles)          # "as in the file", for rebuilding the sheet on save
        sheet.cache.update(cache)
        sheet.loaded_limit = limit + chunk if len(rows) == chunk else None   # None: the whole sheet is in
        self.engine.rescan()
        self.engine.clear()
        if sheet is self.current:
            self.table.reload()
            self.on_select()
        return len(sheet.rows) - first

    def load_all_rows(self, limit=1_048_576):
        """Open the file again with every row (sheets longer than the 'rows to load' setting were cut)."""
        if not self.path or not self.truncated():
            return False
        if not self.confirm_discard():
            return False
        keep = io.MAX_ROWS
        io.MAX_ROWS = limit
        try:
            ok = self.open_file(self.path)
        finally:
            io.MAX_ROWS = keep
        self.status["text"] = f"All rows loaded ({len(self.current.rows):,})" if ok else self.status["text"]
        return ok

    # ---------- keyboard-only ----------

    def post_app_menu(self):
        mb = self.app_menu.master
        self.app_menu.post(mb.winfo_rootx(), mb.winfo_rooty() + mb.winfo_height())
        self.app_menu.focus_set()
        return "break"

    def next_sheet(self, step):
        vis = self.visible_sheets()
        name = vis[(vis.index(self.sheet) + step) % len(vis)]
        self.show(name)
        self.tabs.select(name)
        return "break"

    def keyboard_context_menu(self):
        g = self.table
        r, c = g.cur
        x, y = g.cv.winfo_rootx() + g.col_x(c) + 10, g.cv.winfo_rooty() + g.row_y(r) + 10
        menus.show(self, x, y)
        return "break"

    # ---------- drag & drop cells ----------

    def move_block(self, rect, tr, tc):
        """Move a block's values, formats and comments so its top-left lands on (tr, tc)."""
        g = self.table
        s = g.sheet
        r1, c1, r2, c2 = rect
        dr, dc = tr - r1, tc - c1
        src = {(r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)}
        dst = {(r + dr, c + dc) for r, c in src}
        if any(g.text(r, c) for r, c in dst - src) and not messagebox.askokcancel(
                "Move cells", "There's already data there. Replace it?", parent=self):
            return False
        new = {}
        for r, c in src:
            new[(r, c)] = ("", None, None)
        for r, c in src:
            new[(r + dr, c + dc)] = (g.text(r, c), s.styles.get((r, c)), s.comments.get((r, c)))
        changes = []
        for (r, c), (v, st, note) in new.items():
            changes += [("v", r, c, v), ("s", r, c, st)]
            if note is not None or (r, c) in s.comments:
                changes.append(("cm", r, c, note))
        # formulas anywhere that point into the moved block follow it (Excel's cut & paste)
        texts = {}                                  # (sheet, r, c) -> text after the move
        for name, sh in self.sheets.items():
            for r, row in enumerate(sh.rows):
                for c, t in enumerate(row):
                    if t.startswith("="):
                        texts[(name, r, c)] = t
        for (r, c), (v, _, _) in new.items():
            texts.pop((self.sheet, r, c), None)
            if v.startswith("="):
                texts[(self.sheet, r, c)] = v
        out = {(self.sheet, r, c): v for (r, c), (v, _, _) in new.items()}
        for (name, r, c), t in texts.items():
            nt = formula.move_refs(t, name, self.sheet, rect, dr, dc)
            if nt != t:
                out[(name, r, c)] = nt
        cells = [("v", self.sheets[name], r, c, v) for (name, r, c), v in out.items()]
        cells += [(k, s, r, c, v) for k, r, c, v in changes if k != "v"]
        self.apply_cells(cells)
        g.select(tr, tc, tr + r2 - r1, tc + c2 - c1)
        return True

    # ---------- format painter ----------

    def format_painter(self):
        """Remember the selection's formatting; the next selection made with the mouse gets it."""
        g = self.table
        g.commit()
        r1, c1, r2, c2 = g.sel_cells() if g.multi() else g.span(*g.cur)
        if (r2 - r1 + 1) * (c2 - c1 + 1) > 10000:
            r2, c2 = r1, c1
        self.painter = [[dict(g.sheet.styles.get((r, c)) or {}) for c in range(c1, c2 + 1)] for r in range(r1, r2 + 1)]
        g.cv.configure(cursor="pencil")
        self.status["text"] = "Format painter: select the cells to format (Esc cancels)"
        return self.painter

    def paint_format(self):
        """Apply the remembered formatting to the selection (the pattern repeats over a bigger block)."""
        block, self.painter = getattr(self, "painter", None), None
        self.table.cv.configure(cursor="")
        if not block:
            return False
        h, w = len(block), len(block[0])
        r1 = self.table.sel_cells()[0]
        c1 = self.table.sel_cells()[1]
        self.format_rect(self.table.cells_rects(), lambda st, r, c: dict(block[(r - r1) % h][(c - c1) % w]))
        self.status["text"] = "Formatting copied"
        return True

    def cancel_painter(self):
        if getattr(self, "painter", None):
            self.painter = None
            self.table.cv.configure(cursor="")
            self.status["text"] = ""

    # ---------- P3 tools (each in its own module) ----------

    def data_tool(self, name):
        import data_more
        self.table.commit()
        return getattr(data_more, name)(self)

    def rules_dialog(self, kind):
        import rules_dialog
        self.table.commit()
        return rules_dialog.open_dialog(self, kind)

    def group(self, on):
        import outline
        return outline.group_selection(self, on)

    def insert_chart(self):
        import charts
        self.table.commit()
        return charts.insert_dialog(self)

    def update_links(self):
        self.engine.externals.__self__.refresh() if hasattr(self.engine.externals, "__self__") else None
        self.engine.clear()
        self.table.data_changed()
        self.status["text"] = "Links updated"

    def spell_check(self):
        import spell
        self.table.commit()
        return spell.SpellDialog.open(self)

    def settings_dialog(self):
        import settings_ui
        return settings_ui.SettingsDialog(self)

    def new_from_template(self, path=None):
        import templates
        return templates.new_from_template(self, path)

    def export_html(self, dst=None):
        import html_export
        self.table.commit()
        return html_export.export_dialog(self, dst)

    def show_help(self, what):
        import help_ui
        return help_ui.show(self, what)

    def check_updates(self):
        import updates
        return updates.check(self)

    def split_window(self):
        import split_view
        return split_view.toggle(self)

    def save_extras(self):
        """Callables fn(openpyxl workbook) for things added / removed in the app that aren't cell
        changes: conditional formats, validation rules, charts."""
        import charts
        fns = []
        rich = self.changes("rt")
        if rich:
            def write_rich(wb, rich=rich):
                import richtext
                for (name, r, c), runs in rich.items():
                    cell = wb[name].cell(row=r + 1, column=c + 1)
                    if type(cell).__name__ == "MergedCell":
                        continue
                    text = self.get(self.sheets[name], r, c)
                    cell.value = richtext.to_openpyxl(runs) if runs and richtext.text_of(runs) == text \
                        else io.to_value(text)
            fns.append(write_rich)
        for name, s in self.sheets.items():
            fn = rules.save_fn(name, s)
            if fn:
                fns.append(fn)
            fn = charts.save_fn(name, s)
            if fn:
                fns.append(fn)
            fn = tables.save_fn(name, s)
            if fn:
                fns.append(fn)
            fn = images.save_fn(name, s)
            if fn:
                fns.append(fn)
        return fns

    # ---------- PDF / print ----------

    def export_pdf(self, dst=None, what="This sheet"):
        """The current sheet (or "All sheets" / the "Selection") as a PDF file. Returns the path (None if
        cancelled / failed)."""
        import pdf_export
        import print_preview
        self.table.commit()
        if dst is None:
            named = self.path or getattr(self, "recovered_from", None)
            stem = os.path.splitext(os.path.basename(named))[0] if named else "Untitled"
            dst = filedialog.asksaveasfilename(filetypes=[("PDF", "*.pdf")], defaultextension=".pdf",
                                               initialfile=f"{stem} - {self.sheet}.pdf")
            if not dst:
                return None
        try:
            n = pdf_export.export(self, dst, **print_preview.pdf_args(self, what))
        except Exception as e:
            messagebox.showerror("Export failed", str(e))
            return None
        self.status["text"] = f"Exported {os.path.basename(dst)} ({n} page{'s' if n != 1 else ''})"
        return dst

    def print_sheet(self, what="This sheet"):
        """Print the current sheet: made into a PDF, sent to the printer by the PDF program."""
        import tempfile

        import pdf_export
        dst = os.path.join(tempfile.gettempdir(), f"SimpleSheets-print-{os.getpid()}.pdf")
        if not self.export_pdf(dst, what):
            return None
        how = pdf_export.print_file(dst)
        if how is None:
            messagebox.showerror("Print", f"No program to print PDF files was found.\n\nThe PDF is here:\n{dst}")
        elif how == "opened":
            self.status["text"] = "Opened the PDF - print it from there"
        else:
            self.status["text"] = "Sent to the printer"
        return how

    # ---------- crash recovery ----------

    def autosave_tick(self):
        """Every minute: refresh the recovery copy while there are unsaved changes."""
        try:
            self.autosave()
        finally:
            self.after(recovery.INTERVAL_MS, self.autosave_tick)

    def autosave(self):
        state = (self.hpos, id(self.history[self.hpos - 1]) if self.hpos else None, self.path)
        if not self.dirty():
            if self.autosaved is not None:
                recovery.discard(self.recovery_id)
                self.autosaved = None
            return False
        if state == self.autosaved:
            return False                         # nothing new since the last copy
        try:
            recovery.save(self.recovery_id, self.sheets, self.names, self.path or getattr(self, "recovered_from", None))
        except Exception:
            return False                         # never interrupt work over a recovery copy
        self.autosaved = state
        return True

    def offer_recovery(self):
        """At start: offer the work of windows that closed without saving (crash, power cut…)."""
        for xlsx, meta in recovery.orphans():
            name = os.path.basename(meta.get("source") or "") or "Untitled"
            when = datetime.datetime.fromtimestamp(meta.get("time", 0)).strftime("%d %b %Y %H:%M")
            if not messagebox.askokcancel("Recover unsaved work",
                                          f"Simple Sheets closed before “{name}” was saved.\n\n"
                                          f"Recover the copy from {when}?\n(Cancel deletes it.)"):
                recovery.forget(xlsx)
            elif self.path is None and not self.dirty():
                self.load_recovered(xlsx)
            else:
                self.new_window(xlsx, recover=True)

    def load_recovered(self, xlsx):
        """Show a recovery copy as unsaved work (Save asks where, suggesting the original name)."""
        meta = {}
        try:
            with open(xlsx[:-5] + ".json", encoding="utf-8") as f:
                meta = json.load(f)
        except (OSError, ValueError):
            pass
        if not self.open_file(xlsx):
            return False
        recent.remove(xlsx)
        self.path, self.recovered, self.recovered_from = None, True, meta.get("source")
        self.disk_stamp = None
        recovery.forget(xlsx)
        self.autosave()                          # our own copy right away
        self.refresh()
        self.status["text"] = "Recovered - save it to keep it"
        return True


if __name__ == "__main__":
    if sys.platform == "win32":  # crisp text on scaled displays
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    try:   # a verified update downloaded last time: install it (silently) instead of starting
        import updates
        if getattr(sys, "frozen", False) and updates.apply_pending():
            sys.exit(0)
    except Exception:
        pass
    args = sys.argv[1:]
    script, quit_after = None, "--quit" in args
    if "--run" in args and args.index("--run") + 1 < len(args):   # --run script.py [file] [--quit]
        k = args.index("--run")
        script = args[k + 1]
        del args[k:k + 2]
    args = [a for a in args if a != "--quit"]
    app = App()
    if len(args) > 1 and args[0] == "--recover":
        app.after(50, app.load_recovered, args[1])
    elif args:
        app.after(50, app.open_file, args[0])
    if script:
        import scripting
        app.after(300, lambda: scripting.run_at_start(app, script, quit_after))
    app.after(400, app.offer_recovery)   # work left by a window that crashed
    if not args and not script:            # a plain start (not opening a file): a tip
        import palette
        app.after(900, lambda: palette.first_run(app))
    if settings.get("check_updates") and settings.get("update_url"):
        import updates
        app.after(3000, lambda: updates.check_in_background(app))
    app.mainloop()
