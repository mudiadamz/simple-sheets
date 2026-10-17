"""Scripts that manage the cells.

Tools → Script…: Python code run inside the app with a small API (below). A whole run is one undo step; if
the script fails or is stopped, nothing is changed. print() goes to the output pane.

    sheet["A1"]                      value of a cell (number, text, True/False, or None when empty)
    sheet["A1"] = 5                  set a value; a text starting with "=" is a formula
    sheet["A1:C3"]                   a block: list of rows
    sheet["A1:C3"] = [[1, 2], ...]   fill a block (a single value fills every cell)
    sheet.cell(row, col)             1-based, like Excel (also .cell(row, col, value) to set)
    sheet.formula("B2")              the text as typed ("=SUM(...)")
    sheet.format("A1:C1", bold=True, bg="#ffff00", fg="#ff0000", numfmt="0.00", size=14, italic=True,
                 align="center", wrap=True)
    sheet.last_row, sheet.last_col   size of the data;  sheet.rows()  every row as a list of values
    sheet.name;  selection           the selected block (.address "B2:D9", .rows(), .row / .col = first cell)
    ref(row, col)                    "C2" for ref(2, 3)
    book[name] / book.sheets / book.active

Data → Run command on selection…: any program (PowerShell, Node, R, a .bat, another Python…) gets the
selected cells as CSV on its input; the CSV it prints replaces them (one undo step).

Scripts can do anything Python can do on this computer: only run scripts you trust. Code is never taken
from a workbook; only what you type or open here runs.
"""
import csv
import datetime as dt
import io
import os
import re
import subprocess
import sys
import time
import traceback

from formatting import parse_number
from formula import XLError, col_index, col_letters

STYLE_KEYS = {"bold", "italic", "underline", "strike", "fg", "bg", "size", "font", "numfmt", "align", "valign",
              "wrap", "indent", "rotate", "border", "border_style", "border_color"}
BORDER_STYLES = {"thin", "medium", "thick", "dashed", "dotted", "double", "hair"}


def _lines(spec, axis):
    """"B", "B:D", 3 (1-based), "3:5", ["A", "C"] -> 0-based indexes."""
    if isinstance(spec, (list, tuple)):
        return [i for s in spec for i in _lines(s, axis)]
    if isinstance(spec, int):
        return [spec - 1]
    spec = str(spec).replace("$", "").strip().upper()
    a, _, b = spec.partition(":")
    b = b or a
    if axis == "col" and a.isalpha() and b.isalpha():
        i, j = col_index(a), col_index(b)
    elif a.isdigit() and b.isdigit():
        i, j = int(a) - 1, int(b) - 1
    else:
        raise ScriptError(f"Not a {'column' if axis == 'col' else 'row'}: {spec!r}")
    return list(range(min(i, j), max(i, j) + 1))


class ScriptError(Exception):
    pass


class Stopped(BaseException):          # not an Exception: a script's own `except Exception` can't swallow it
    pass


def parse_area(ref):
    """'A1' / 'B2:D9' / 'C:C' / '3:5' -> (r1, c1, r2, c2), 0-based (whole columns / rows end at the data)."""
    ref = ref.replace("$", "").strip().upper()
    m = re.fullmatch(r"([A-Z]{1,3})(\d+)(?::([A-Z]{1,3})(\d+))?", ref)
    if m:
        r1, c1 = int(m.group(2)) - 1, col_index(m.group(1))
        r2, c2 = (int(m.group(4)) - 1, col_index(m.group(3))) if m.group(3) else (r1, c1)
        return min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)
    m = re.fullmatch(r"([A-Z]{1,3}):([A-Z]{1,3})", ref)
    if m:
        return 0, col_index(m.group(1)), None, col_index(m.group(2))
    m = re.fullmatch(r"(\d+):(\d+)", ref)
    if m:
        return int(m.group(1)) - 1, 0, int(m.group(2)) - 1, None
    raise ScriptError(f"Not a cell or range: {ref!r}")


def to_text(v):
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (int, float)):
        return str(int(v)) if float(v).is_integer() and abs(v) < 1e15 else repr(float(v))
    if isinstance(v, dt.datetime):
        return v.strftime("%Y-%m-%d %H:%M:%S") if v.time() != dt.time() else v.strftime("%Y-%m-%d")
    if isinstance(v, dt.date):
        return v.strftime("%Y-%m-%d")
    return str(v)


def to_value(v):
    """What a script reads: numbers as int / float, empty as None, TRUE/FALSE as bool, errors as text."""
    if v is None or v == "":
        return None
    if isinstance(v, bool):
        return v
    if isinstance(v, XLError):
        return str(v)
    if isinstance(v, (int, float)):
        return int(v) if float(v).is_integer() and abs(v) < 1e15 else float(v)
    if isinstance(v, str):
        if v in ("TRUE", "FALSE"):
            return v == "TRUE"
        n = parse_number(v)
        if n is not None:
            return int(n) if float(n).is_integer() and abs(n) < 1e15 else n
    return v


class Run:
    """One script run: changes go into the model right away (so reads see them, formulas included) and the
    originals are kept, to roll back or to turn into one undo step at the end. Any undoable kind of change
    can be recorded (values "v", formats "s", freeze "fz", merges "m", widths "w", ... see App.read); inserting /
    deleting rows or columns and adding sheets are 'structural' and use a whole-workbook snapshot."""

    # kinds the whole-workbook snapshot (structure.FIELDS) already covers
    IN_SNAPSHOT = {"v", "s", "w", "rh", "m", "fz", "flt", "cm", "ln", "hr", "hc", "or", "oc", "rt", "th"}

    def __init__(self, app):
        import structure
        self.app = app
        self.orig = {}                     # (kind, Sheet, r, c) -> value before the run
        self.dirty = False
        self.structural = self.book_changed = False
        self.snap = structure.snapshot(app.sheets)
        self.order = list(app.sheets.items())
        self.names = dict(app.names)
        self.sheet_before = app.sheet

    def name_of(self, s):
        return self.app.name_of(s)

    def check(self, s, kind, r, c):
        if s is not None and s.props.get("protect"):
            from workbook_tools import is_locked
            if kind != "v" or is_locked(self.app, s, r, c):
                what = f"{col_letters(c)}{r + 1}" if kind in ("v", "s") else "its layout"
                raise ScriptError(f"{self.name_of(s)} is protected: {what} can't be changed")

    def write(self, kind, s, r, c, val):
        self.check(s, kind, r, c)
        key = (kind, s, r, c)
        if key not in self.orig:
            self.orig[key] = self.app.read(kind, s, r, c)
        if kind == "v":
            while len(s.rows) <= r:
                s.rows.append([])
            row = s.rows[r]
            while len(row) <= c:
                row.append("")
            row[c] = val
        elif kind == "s":
            if val:
                s.styles[(r, c)] = val
            else:
                s.styles.pop((r, c), None)
        else:
            self.app.write_cell(kind, s, r, c, val)
        self.dirty = True

    def structure(self, fn):
        """A change of rows / columns / sheets: done now, undone from the snapshot."""
        self.structural = True
        self.dirty = True
        return fn()

    def value(self, s, r, c):
        if self.dirty:
            self.app.engine.clear()
            self.dirty = False
        t = s.rows[r][c] if r < len(s.rows) and c < len(s.rows[r]) else ""
        if t.startswith("=") and len(t) > 1:
            return to_value(self.app.engine.cell(self.name_of(s), r, c))
        return to_value(t)

    def rollback(self):
        import structure
        for (kind, s, r, c), old in self.orig.items():
            if kind == "v":
                s.rows[r][c] = old
            elif kind == "s":
                if old:
                    s.styles[(r, c)] = old
                else:
                    s.styles.pop((r, c), None)
            else:
                self.app.write_cell(kind, s, r, c, old)
        if self.structural or self.book_changed:
            self.app.sheets.clear()
            self.app.sheets.update(self.order)
            structure.restore(self.app.sheets, self.snap)
            self.app.names = dict(self.names)
            self.app.engine.set_names(self.app.names)
            if self.app.sheet not in self.app.sheets:
                self.app.sheet = self.sheet_before
            self.app.table.load(self.app.current)
        self.app.engine.clear()

    def finish(self):
        """Back to the originals, then every change again through the app: one undo step. -> things changed."""
        import structure
        final = {key: self.app.read(*key) for key in self.orig}
        changed = [key for key in self.orig if final[key] != self.orig[key]]
        if not (self.structural or self.book_changed):
            self.rollback()
            if changed:
                self.app.apply_cells([key + (final[key],) for key in changed])
            return len({(s, r, c) for _k, s, r, c in changed})
        # structural: the snapshot entry restores rows / formats / sizes / sheets; the other kinds go alongside
        g = self.app.table
        if self.book_changed:
            before = (self.order, self.snap, self.sheet_before, self.names)
            group = [("b", None, 0, 0, before, self.app.book_state(self.app.sheet))]
        else:
            group = [("x", self.app.current, g.cur[0], g.cur[1], self.snap, structure.snapshot(self.app.sheets))]
        for key in changed:
            kind, s, r, c = key
            self.app.base.setdefault(key, self.orig[key])          # so saving writes them
            if kind not in self.IN_SNAPSHOT and not (self.book_changed and kind == "nm"):
                group.append(key + (self.orig[key], final[key]))
        if self.app.names != self.names and not self.book_changed:
            group.append(("nm", None, 0, 0, dict(self.names), dict(self.app.names)))
        self.app.push(group)
        return max(len(changed), 1)


class Block:
    """A range (also the `selection`)."""

    def __init__(self, sheet, r1, c1, r2, c2):
        self.sheet, self.r1, self.c1, self.r2, self.c2 = sheet, r1, c1, r2, c2

    @property
    def row(self):
        """First row, 1-based like Excel."""
        return self.r1 + 1

    @property
    def col(self):
        return self.c1 + 1

    @property
    def address(self):
        a = f"{col_letters(self.c1)}{self.r1 + 1}"
        return a if (self.r1, self.c1) == (self.r2, self.c2) else f"{a}:{col_letters(self.c2)}{self.r2 + 1}"

    def rows(self):
        return [[self.sheet._get(r, c) for c in range(self.c1, self.c2 + 1)] for r in range(self.r1, self.r2 + 1)]

    @property
    def values(self):
        return self.rows()

    @values.setter
    def values(self, data):
        self.sheet[self.address] = data

    def __iter__(self):
        return iter(self.rows())

    def __repr__(self):
        return f"<{self.sheet.name}!{self.address}>"


class Sheet:
    def __init__(self, run, s):
        self._run, self._s = run, s

    @property
    def name(self):
        return self._run.name_of(self._s)

    @property
    def last_row(self):
        rows = self._s.rows
        return max((r + 1 for r, row in enumerate(rows) if any(t != "" for t in row)), default=0)

    @property
    def last_col(self):
        return max((max((c + 1 for c, t in enumerate(row) if t != ""), default=0) for row in self._s.rows), default=0)

    def _area(self, ref):
        r1, c1, r2, c2 = parse_area(ref)
        return r1, c1, max(self.last_row - 1, r1) if r2 is None else r2, max(self.last_col - 1, c1) if c2 is None else c2

    def _get(self, r, c):
        return self._run.value(self._s, r, c)

    def __getitem__(self, ref):
        r1, c1, r2, c2 = self._area(ref)
        if (r1, c1) == (r2, c2) and ":" not in ref:
            return self._get(r1, c1)
        return Block(self, r1, c1, r2, c2).rows()

    def __setitem__(self, ref, value):
        r1, c1, r2, c2 = self._area(ref)
        if isinstance(value, (list, tuple)):
            rows = [list(v) if isinstance(v, (list, tuple)) else [v] for v in value]
            if ":" not in ref:                 # one cell given rows of values: they spread from there
                r2, c2 = r1 + len(rows) - 1, c1 + max((len(x) for x in rows), default=1) - 1
            for i, row in enumerate(rows):
                for j, v in enumerate(row):
                    if r1 + i <= r2 and c1 + j <= c2:
                        self._run.write("v", self._s, r1 + i, c1 + j, to_text(v))
        else:
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    self._run.write("v", self._s, r, c, to_text(value))

    def cell(self, row, col, *value):
        """1-based row / column; with a third argument, sets the cell."""
        if value:
            self._run.write("v", self._s, row - 1, col - 1, to_text(value[0]))
            return value[0]
        return self._get(row - 1, col - 1)

    def clear(self, ref=None, formats=False):
        """Empty cells (the whole used area when no ref); formats=True also removes their formatting."""
        if ref is None:
            r1, c1, r2, c2 = 0, 0, self.last_row - 1, self.last_col - 1
        else:
            r1, c1, r2, c2 = self._area(ref)
        s = self._s
        for r in range(r1, min(r2, len(s.rows) - 1) + 1):
            for c in range(c1, min(c2, len(s.rows[r]) - 1) + 1):
                if s.rows[r][c] != "":
                    self._run.write("v", s, r, c, "")
        if formats:
            for (r, c) in [k for k in s.styles if r1 <= k[0] <= r2 and c1 <= k[1] <= c2]:
                self._run.write("s", s, r, c, None)

    def formula(self, ref):
        r, c, _, _ = self._area(ref)
        s = self._s
        return s.rows[r][c] if r < len(s.rows) and c < len(s.rows[r]) else ""

    def format(self, ref, **style):
        bad = set(style) - STYLE_KEYS
        if bad:
            raise ScriptError(f"Unknown format: {', '.join(sorted(bad))} (use {', '.join(sorted(STYLE_KEYS))})")
        r1, c1, r2, c2 = self._area(ref)
        border = style.pop("border", None)
        bstyle, bcolor = style.pop("border_style", "thin"), style.pop("border_color", "#000000")
        if bstyle not in BORDER_STYLES:
            raise ScriptError(f"border_style: one of {', '.join(sorted(BORDER_STYLES))}")
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                st = dict(self._s.styles.get((r, c)) or {})
                for k, v in style.items():
                    if v in (None, False):
                        st.pop(k, None)
                    else:
                        st[k] = v
                if border is not None:
                    st = _with_border(st, border, (r, c), (r1, c1, r2, c2), bstyle, bcolor)
                self._run.write("s", self._s, r, c, st or None)

    def range(self, ref):
        return Block(self, *self._area(ref))

    # ---- layout: freeze, merge, sizes, hiding ----

    def freeze(self, rows=0, cols=0):
        """freeze(1) = the top row, freeze(cols=1) = column A, freeze("B2") = rows above and columns left of B2."""
        if isinstance(rows, str):
            r, c, _, _ = self._area(rows)
            rows, cols = r, c
        self._run.write("fz", self._s, 0, 0, (int(rows), int(cols)))

    def unfreeze(self):
        self._run.write("fz", self._s, 0, 0, (0, 0))

    def merge(self, ref):
        r1, c1, r2, c2 = self._area(ref)
        keep = [m for m in self._s.merges if m[2] < r1 or m[0] > r2 or m[3] < c1 or m[1] > c2]
        self._run.write("m", self._s, 0, 0, keep + [(r1, c1, r2, c2)])

    def unmerge(self, ref):
        r1, c1, r2, c2 = self._area(ref)
        self._run.write("m", self._s, 0, 0, [m for m in self._s.merges
                                             if m[2] < r1 or m[0] > r2 or m[3] < c1 or m[1] > c2])

    def width(self, cols, chars):
        """Column width in characters, e.g. width("B", 20) or width("B:D", 12)."""
        for c in _lines(cols, "col"):
            self._run.write("w", self._s, 0, c, float(chars) if chars else None)

    def autofit(self, cols=None):
        """Columns as wide as their longest text (all used columns when none are given)."""
        cols = _lines(cols, "col") if cols is not None else range(self.last_col)
        for c in cols:
            longest = max((len(str(self._get(r, c) if self._get(r, c) is not None else ""))
                           for r in range(self.last_row)), default=0)
            self._run.write("w", self._s, 0, c, float(min(max(longest * 1.15 + 2, 6), 80)))

    def height(self, rows, points):
        """Row height in points, e.g. height(1, 30) or height("2:5", 18)."""
        for r in _lines(rows, "row"):
            self._run.write("rh", self._s, r, 0, float(points) if points else None)

    def hide_rows(self, rows, hide=True):
        for r in _lines(rows, "row"):
            self._run.write("hr", self._s, r, 0, hide)

    def show_rows(self, rows):
        self.hide_rows(rows, False)

    def hide_cols(self, cols, hide=True):
        for c in _lines(cols, "col"):
            self._run.write("hc", self._s, 0, c, hide)

    def show_cols(self, cols):
        self.hide_cols(cols, False)

    # ---- insert / delete / sort ----

    def _restructure(self, axis, at, n):
        import formula as fm
        import structure
        run, app, name = self._run, self._run.app, self.name
        run.check(self._s, "x", 0, 0)

        def do():
            for home, r, c in structure.change(app.sheets, name, axis, at, n):
                if home != name:                  # formulas on other sheets: saved as cell edits
                    before = run.snap.get(home, {}).get("rows", [])
                    old = before[r][c] if r < len(before) and c < len(before[r]) else ""
                    app.base.setdefault(("v", app.sheets[home], r, c), old)
            app.names = {k: fm.adjust(v, name, name, axis, at, n) for k, v in app.names.items()}
            app.engine.set_names(app.names)
        run.structure(do)

    def insert_rows(self, before, count=1):
        """New empty rows above row `before` (1-based)."""
        self._restructure("row", int(before) - 1, int(count))

    def delete_rows(self, first, count=1):
        self._restructure("row", int(first) - 1, -int(count))

    def insert_cols(self, before, count=1):
        """New empty columns left of column `before` ("C" or 3)."""
        self._restructure("col", _lines(before, "col")[0], int(count))

    def delete_cols(self, first, count=1):
        self._restructure("col", _lines(first, "col")[0], -int(count))

    def sort(self, ref, by, descending=False, header=True):
        """Sort the rows of a block by a column ("B") or several (["B", "A"]); formats move with the rows and
        formulas are adjusted like Excel does."""
        import formula as fm
        r1, c1, r2, c2 = self._area(ref)
        first = r1 + 1 if header else r1
        keys = [_lines(k, "col")[0] for k in (by if isinstance(by, (list, tuple)) else [by])]

        def key(r):
            out = []
            for c in keys:
                v = self._get(r, c)
                blank = v is None
                num = isinstance(v, (int, float)) and not isinstance(v, bool)
                out.append((blank, 0 if num else 1, v if num else str(v).casefold() if v is not None else ""))
            return out
        order = sorted(range(first, r2 + 1), key=key, reverse=descending)
        if descending:                        # empty cells stay at the end either way
            order = [r for r in order if key(r)[0][0] is False] + [r for r in order if key(r)[0][0] is True]
        s = self._s
        texts = {(r, c): (s.rows[r][c] if r < len(s.rows) and c < len(s.rows[r]) else "")
                 for r in range(first, r2 + 1) for c in range(c1, c2 + 1)}
        styles = {(r, c): s.styles.get((r, c)) for r in range(first, r2 + 1) for c in range(c1, c2 + 1)}
        for new_r, old_r in zip(range(first, r2 + 1), order):
            for c in range(c1, c2 + 1):
                t = texts[(old_r, c)]
                if t.startswith("=") and new_r != old_r:
                    t = "=" + fm.shift(t[1:], new_r - old_r, 0)
                self._run.write("v", s, new_r, c, t)
                self._run.write("s", s, new_r, c, styles[(old_r, c)])

    # ---- filter, comments, links, tab colour ----

    def filter(self, ref=None, column=None, values=None):
        """Filter buttons on a block (the data around A1 when no ref); with column= and values=, show only those
        rows: filter("A1:E40", column="E", values=["Delivered"])."""
        if ref is None:
            r1, c1 = 0, 0
            r2, c2 = max(self.last_row - 1, 0), max(self.last_col - 1, 0)
        else:
            r1, c1, r2, c2 = self._area(ref)
        crit = {}
        if column is not None:
            crit[_lines(column, "col")[0]] = [to_text(v) for v in (values or [])]
        self._run.write("flt", self._s, 0, 0, {"range": (r1, c1, r2, c2), "criteria": crit})

    def clear_filter(self):
        self._run.write("flt", self._s, 0, 0, None)

    def comment(self, ref, text):
        r, c, _, _ = self._area(ref)
        self._run.write("cm", self._s, r, c, [str(text), "Script"] if text else None)

    def link(self, ref, url):
        r, c, _, _ = self._area(ref)
        self._run.write("ln", self._s, r, c, str(url) if url else None)

    def tab_color(self, color):
        self._run.write("sp", self._s, 0, "tab", color or None)

    def on_click(self, ref, code):
        """Clicking the cell runs this code (a button in the sheet): on_click("G2", "sheet['B2'] = 2026").
        The code gets `clicked` (its cell's address); on_click("G2", None) removes it."""
        import cell_scripts
        app = self._run.app
        r, c, _, _ = self._area(ref)
        name = cell_scripts.sheet_name(app, self._s)
        self._run.write("wp", None, 0, "scripts", cell_scripts.changed(app, name, cell_scripts.ref_of(r, c), code))
        cell_scripts.trust(app)                  # made here, so it may run without asking

    # ---- drop-down lists, conditional formatting ----

    def dropdown(self, ref, items):
        """A drop-down list of choices in the cells: dropdown("D2:D50", ["Yes", "No"])."""
        rect = self._area(ref)
        f1 = '"' + ",".join(str(i).replace('"', "") for i in items) + '"'
        rule = {"ranges": [rect], "model": True, "type": "list", "op": "between", "f1": f1, "f2": None, "arrow": True,
                "show_error": True, "style": "stop", "title": "", "error": ""}
        self._run.write("dv", self._s, 0, 0, [rule] + list(self._s.validations))

    CF_OPS = {">": "greaterThan", "<": "lessThan", ">=": "greaterThanOrEqual", "<=": "lessThanOrEqual",
              "=": "equal", "==": "equal", "!=": "notEqual", "<>": "notEqual", "between": "between"}

    def highlight(self, ref, rule, value=None, value2=None, **style):
        """Conditional formatting: highlight("C2:C40", ">", 100, bg="#ffc7ce", fg="#9c0006"). rule: > < >= <= =
        != between, contains, begins, ends, top (value = how many), bottom, above_average, below_average,
        formula (value = "=$C2>$D2")."""
        import rules_dialog
        bad = set(style) - {"bg", "fg", "bold", "italic", "underline", "strike"}
        if bad:
            raise ScriptError(f"highlight: unknown format {', '.join(bad)} (use bg, fg, bold, italic)")
        style = style or {"bg": "#ffc7ce", "fg": "#9c0006"}
        rect = [self._area(ref)]
        v1, v2 = to_text(value) if value is not None else "", to_text(value2) if value2 is not None else ""
        if rule in self.CF_OPS:
            d = rules_dialog.make_cf("cellIs", rect, self.CF_OPS[rule], v1, v2, style)
        else:
            kind = {"contains": "containsText", "begins": "beginsWith", "ends": "endsWith", "top": "top",
                    "bottom": "bottom", "above_average": "above", "below_average": "below",
                    "formula": "expression", "blank": "containsBlanks"}.get(rule)
            if kind is None:
                raise ScriptError(f"highlight: unknown rule {rule!r}")
            d = rules_dialog.make_cf(kind, rect, None, v1, v2, style)
        self._run.write("cf", self._s, 0, 0, [d] + list(self._s.cond))

    def color_scale(self, ref, colors=("#f8696b", "#ffeb84", "#63be7b")):
        """Colours from the lowest to the highest value (2 or 3 colours)."""
        import rules_dialog
        d = rules_dialog.make_cf("scale3" if len(colors) == 3 else "scale2", [self._area(ref)])
        pts = d["scale"]
        d["scale"] = [(p[0], p[1], col) for p, col in zip(pts, colors)]
        self._run.write("cf", self._s, 0, 0, [d] + list(self._s.cond))

    def data_bar(self, ref, color="#638ec6"):
        import rules_dialog
        d = rules_dialog.make_cf("dataBar", [self._area(ref)])
        d["bar"] = (d["bar"][0], color)
        self._run.write("cf", self._s, 0, 0, [d] + list(self._s.cond))

    def clear_rules(self, ref=None):
        """Remove conditional formatting and drop-down lists (all, or those overlapping ref)."""
        def keep(rules):
            if ref is None:
                return []
            r1, c1, r2, c2 = self._area(ref)
            return [d for d in rules if all(b[2] < r1 or b[0] > r2 or b[3] < c1 or b[1] > c2 for b in d["ranges"])]
        self._run.write("cf", self._s, 0, 0, keep(self._s.cond))
        self._run.write("dv", self._s, 0, 0, keep(self._s.validations))


    def rows(self):
        return Block(self, 0, 0, max(self.last_row - 1, 0), max(self.last_col - 1, 0)).rows() if self.last_row else []

    def __repr__(self):
        return f"<sheet {self.name!r}>"


def _with_border(st, border, cell, rect, style, color):
    """border: "all", "outline", "inside"... or "top" / "bottom" / "left" / "right" / "none", or a list of those."""
    r, c = cell
    r1, c1, r2, c2 = rect
    sides = dict(st.get("border") or {})
    wanted = set()
    for b in ([border] if isinstance(border, str) else list(border or [])):
        if b == "none":
            sides = {}
        elif b == "all":
            wanted |= {"top", "bottom", "left", "right"}
        elif b == "outline":
            wanted |= {s for s, on in (("top", r == r1), ("bottom", r == r2), ("left", c == c1), ("right", c == c2)) if on}
        elif b in ("top", "bottom", "left", "right"):
            wanted.add(b)
        else:
            raise ScriptError(f"border: all, outline, top, bottom, left, right or none - not {b!r}")
    for side in wanted:
        sides[side] = [style, color]
    out = dict(st)
    if sides:
        out["border"] = sides
    else:
        out.pop("border", None)
    return out


class Book:
    def __init__(self, run):
        self._run = run

    def __getitem__(self, name):
        if name not in self._run.app.sheets:
            raise ScriptError(f"No sheet named {name!r} (sheets: {', '.join(self._run.app.sheets)})")
        return Sheet(self._run, self._run.app.sheets[name])

    @property
    def sheets(self):
        return [self[n] for n in self._run.app.sheets]

    @property
    def active(self):
        return self[self._run.app.sheet]

    def save(self):
        """Saved after the script has finished (with its changes)."""
        self._run.save_after = True

    def add_sheet(self, name=None):
        """A new empty sheet after the active one (a free name when none is given) -> that sheet."""
        import sheet_io as io
        app, run = self._run.app, self._run
        if name is None:
            n = len(app.sheets) + 1
            while app.check_name(f"Sheet{n}"):
                n += 1
            name = f"Sheet{n}"
        err = app.check_name(name)
        if err:
            raise ScriptError(err)
        run.book_changed = True
        items = list(app.sheets.items())
        at = list(app.sheets).index(app.sheet) + 1 if app.sheet in app.sheets else len(items)
        items.insert(at, (name, io.Sheet([])))
        app.sheets.clear()
        app.sheets.update(items)
        return self[name]

    def rename_sheet(self, old, new):
        import formula as fm
        app, run = self._run.app, self._run
        if old not in app.sheets:
            raise ScriptError(f"No sheet named {old!r}")
        err = app.check_name(new, old)
        if err:
            raise ScriptError(err)
        run.check(app.sheets[old], "x", 0, 0)
        run.book_changed = True
        items = [(new if n == old else n, s) for n, s in app.sheets.items()]
        app.sheets.clear()
        app.sheets.update(items)
        for sh in app.sheets.values():
            for row in sh.rows:
                for c, t in enumerate(row):
                    if t.startswith("="):
                        row[c] = fm.rename_sheet(t, old, new)
        app.names = {k: fm.rename_sheet(v, old, new) for k, v in app.names.items()}
        if app.sheet == old:
            app.sheet = new
        return self[new]


def run(app, code, out=None, stop=None, timeout=None, filename="<script>", extra=None):
    """Run a script. out(text) gets print() output; stop() -> True asks it to stop. Returns (ok, message)."""
    app.table.commit()
    r = Run(app)
    r.save_after = False
    sheet = Sheet(r, app.current)
    g = app.table
    sel = Block(sheet, *g.sel_cells())
    lines = []

    def emit(*args, sep=" ", end="\n", **_kw):
        text = sep.join(str(a) for a in args) + end
        lines.append(text)
        if out:
            out(text)

    env = {"__name__": "__script__", "book": Book(r), "sheet": sheet, "selection": sel, "print": emit,
           "ref": lambda row, col: f"{col_letters(col - 1)}{row}",          # ref(2, 3) -> "C2"
           "app_version": __import__("version").APP_VERSION}
    env.update(extra or {})                  # e.g. the AI assistant's attached files
    started = time.monotonic()
    count = [0]

    def tracer(frame, event, arg):
        count[0] += 1
        if count[0] % 2000 == 0:
            if (stop and stop()) or (timeout and time.monotonic() - started > timeout):
                raise Stopped()
        return tracer
    old = sys.gettrace()
    sys.settrace(tracer)
    try:
        exec(compile(code, filename, "exec"), env)
    except Stopped:
        r.rollback()
        return False, "Stopped: nothing was changed."
    except Exception as e:
        r.rollback()
        tb = traceback.extract_tb(e.__traceback__)
        line = next((f.lineno for f in reversed(tb) if f.filename == filename), None)
        where = f" (line {line})" if line else ""
        return False, f"{type(e).__name__}{where}: {e}\nNothing was changed."
    finally:
        sys.settrace(old)
    n = r.finish()
    if r.save_after:
        app.save()
    return True, f"Done: {n} cell{'s' if n != 1 else ''} changed." if n else "Done: no cells changed."


# ---------- the script window ----------

EXAMPLE = '''# Python script: `sheet` is the active sheet, `selection` the selected cells, `book` the workbook.
# Everything it changes is one undo step (Ctrl+Z). Press F5 to run.

# Double every number in the selection:
for i, row in enumerate(selection.rows()):
    for j, v in enumerate(row):
        if isinstance(v, (int, float)):
            sheet.cell(selection.row + i, selection.col + j, v * 2)

# A few more things you can do:
# sheet["A1"] = "Total"
# sheet["B1"] = "=SUM(B2:B100)"
# sheet.format("A1:B1", bold=True, bg="#ffff00")
# print(sheet.last_row, "rows", sheet["A2:B5"])
'''

SNIPPETS = {
    "Double the numbers in the selection": EXAMPLE,
    "Number the rows (1, 2, 3…) in column A": 'for r in range(2, sheet.last_row + 1):\n    sheet.cell(r, 1, r - 1)\n',
    "Upper-case the selected text": ('for i, row in enumerate(selection.rows()):\n    for j, v in enumerate(row):\n'
                                     '        if isinstance(v, str):\n'
                                     '            sheet.cell(selection.row + i, selection.col + j, v.upper())\n'),
    "Highlight negative numbers in red": ('for r, row in enumerate(sheet.rows(), start=1):\n'
                                          '    for c, v in enumerate(row, start=1):\n'
                                          '        if isinstance(v, (int, float)) and v < 0:\n'
                                          '            sheet.format(ref(r, c), fg="#d93025", bold=True)\n'),
    "List every sheet and its size": 'for s in book.sheets:\n    print(s.name, s.last_row, "rows x", s.last_col, "columns")\n',
}


def scripts_dir():
    from simplecore import recent
    d = os.path.join(recent.config_dir(), "scripts")
    os.makedirs(d, exist_ok=True)
    return d


class ScriptWindow:
    def __init__(self, app):
        import tkinter as tk
        from tkinter import ttk

        from simplecore.theme import C
        self.app, self.path, self.running, self.stop_flag = app, None, False, False
        w = self.win = tk.Toplevel(app)
        w.title("Script")
        w.transient(app)
        w.geometry(f"{int(720 * app.scale)}x{int(560 * app.scale)}")
        bar = ttk.Frame(w, padding=(6, 6, 6, 0))
        bar.pack(fill="x")
        self.run_btn = ttk.Button(bar, text="▶ Run (F5)", command=self.run)
        self.run_btn.pack(side="left")
        self.stop_btn = ttk.Button(bar, text="■ Stop", command=self.stop, state="disabled")
        self.stop_btn.pack(side="left", padx=4)
        ttk.Button(bar, text="Open…", command=self.open).pack(side="left", padx=(12, 0))
        ttk.Button(bar, text="Save…", command=self.save).pack(side="left", padx=4)
        ttk.Button(bar, text="Guide", command=lambda: app.show_help("scripts")).pack(side="left", padx=(12, 0))
        self.snippet = ttk.Combobox(bar, values=list(SNIPPETS), state="readonly", width=36)
        self.snippet.set("Examples…")
        self.snippet.pack(side="right")
        self.snippet.bind("<<ComboboxSelected>>", lambda e: self.load_text(SNIPPETS[self.snippet.get()]))
        mono = ("Consolas", 10)
        pane = ttk.PanedWindow(w, orient="vertical")
        pane.pack(fill="both", expand=True, padx=6, pady=6)
        self.code = tk.Text(pane, wrap="none", undo=True, font=mono, background=C["field_bg"], foreground=C["fg"],
                            insertbackground=C["fg"], tabs=("4c",))
        self.output = tk.Text(pane, height=7, wrap="word", font=mono, state="disabled", background=C["bg"],
                              foreground=C["fg"])
        pane.add(self.code, weight=4)
        pane.add(self.output, weight=1)
        ttk.Label(w, text="Scripts can do anything Python can on this computer - only run code you trust.",
                  foreground=C["muted"]).pack(anchor="w", padx=8, pady=(0, 6))
        self.code.bind("<Tab>", lambda e: (self.code.insert("insert", "    "), "break")[-1])
        w.bind("<F5>", lambda e: (self.run(), "break")[-1])
        w.bind("<Escape>", lambda e: self.stop() if self.running else w.destroy())
        last = os.path.join(scripts_dir(), "last.py")
        try:
            with open(last, encoding="utf-8") as f:
                text = f.read()
        except OSError:
            text = EXAMPLE
        self.load_text(text or EXAMPLE)
        self.code.focus_set()

    def load_text(self, text):
        self.code.delete("1.0", "end")
        self.code.insert("1.0", text)

    def say(self, text, clear=False):
        self.output.configure(state="normal")
        if clear:
            self.output.delete("1.0", "end")
        self.output.insert("end", text)
        self.output.see("end")
        self.output.configure(state="disabled")
        self.output.update_idletasks()

    def stop(self):
        self.stop_flag = True

    def stopping(self):
        self.win.update()                   # keeps the window (and its Stop button) alive while running
        return self.stop_flag

    def run(self):
        if self.running:
            return None
        code = self.code.get("1.0", "end-1c")
        try:
            with open(os.path.join(scripts_dir(), "last.py"), "w", encoding="utf-8") as f:
                f.write(code)
        except OSError:
            pass
        self.running, self.stop_flag = True, False
        self.run_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.say("", clear=True)
        self.win.grab_set()                 # the sheet waits while the script runs
        try:
            ok, msg = run(self.app, code, out=self.say, stop=self.stopping, filename=self.path or "<script>")
        finally:
            self.win.grab_release()
            self.running = False
            if self.win.winfo_exists():
                self.run_btn.configure(state="normal")
                self.stop_btn.configure(state="disabled")
        if self.win.winfo_exists():
            self.say(("" if ok else "\n") + msg + "\n")
        self.app.status["text"] = "Script: " + msg.split("\n")[0]
        return ok

    def open(self, path=None):
        from tkinter import filedialog
        path = path or filedialog.askopenfilename(parent=self.win, initialdir=scripts_dir(),
                                                  filetypes=[("Python scripts", "*.py"), ("All files", "*.*")])
        if path:
            with open(path, encoding="utf-8") as f:
                self.load_text(f.read())
            self.path = path
            self.win.title(f"Script - {os.path.basename(path)}")
        return path

    def save(self, path=None):
        from tkinter import filedialog
        path = path or filedialog.asksaveasfilename(parent=self.win, initialdir=scripts_dir(), defaultextension=".py",
                                                    filetypes=[("Python scripts", "*.py")])
        if path:
            with open(path, "w", encoding="utf-8") as f:
                f.write(self.code.get("1.0", "end-1c"))
            self.path = path
            self.win.title(f"Script - {os.path.basename(path)}")
        return path


def open_window(app):
    return ScriptWindow(app)


def run_at_start(app, path, quit_after=False):
    """SimpleSheets --run script.py [file] [--quit]: run a script on the opened file (automation). The result
    goes to the status bar and to <script>.log; with --quit the app closes afterwards (exit code 1 on failure;
    a script that wants its changes kept calls book.save())."""
    try:
        with open(path, encoding="utf-8") as f:
            code = f.read()
        ok, msg = run(app, code, filename=path)
    except OSError as e:
        ok, msg = False, f"Can't read the script: {e}"
    app.status["text"] = "Script: " + msg.split("\n")[0]
    try:
        with open(os.path.splitext(path)[0] + ".log", "a", encoding="utf-8") as f:
            f.write(f"{dt.datetime.now():%Y-%m-%d %H:%M:%S} {'OK' if ok else 'FAILED'} {msg}\n")
    except OSError:
        pass
    if quit_after:
        app.destroy()
        sys.exit(0 if ok else 1)
    return ok, msg


# ---------- any program, through CSV ----------

def run_command(app, command, timeout=60):
    """Send the selected cells as CSV to `command`'s input; the CSV it prints replaces them (and may be
    bigger: it fills from the selection's top-left). One undo step. -> (ok, message)."""
    g = app.table
    g.commit()
    r1, c1, r2, c2 = g.sel_cells()
    rows = [[g.value(r, c) for c in range(c1, c2 + 1)] for r in range(r1, r2 + 1)]
    buf = io.StringIO()
    csv.writer(buf, lineterminator="\n").writerows(rows)
    try:
        p = subprocess.run(command, input=buf.getvalue(), capture_output=True, text=True, shell=True,
                           timeout=timeout, encoding="utf-8", errors="replace",
                           creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except subprocess.TimeoutExpired:
        return False, f"The command took longer than {timeout} s and was stopped. Nothing was changed."
    if p.returncode != 0:
        return False, f"The command failed (exit code {p.returncode}):\n{(p.stderr or p.stdout).strip()[:2000]}"
    new = list(csv.reader(io.StringIO(p.stdout)))
    if not new:
        return False, "The command printed nothing. Nothing was changed."
    changes = []
    for i, row in enumerate(new):
        for j, v in enumerate(row):
            changes.append((r1 + i, c1 + j, v))
    for r in range(r1, r2 + 1):                          # selected cells it didn't print: emptied
        for c in range(c1, c2 + 1):
            if r - r1 >= len(new) or c - c1 >= len(new[r - r1]):
                changes.append((r, c, ""))
    app.set_cells(changes)
    return True, f"Done: {len(new)} row{'s' if len(new) != 1 else ''} back from the command."


def run_command_dialog(app):
    from tkinter import messagebox, simpledialog

    from simplecore import settings
    cmd = simpledialog.askstring(
        "Run command on selection",
        "A command that reads CSV on its input and prints CSV, e.g.\n"
        "   python my_script.py\n   powershell -File clean.ps1\n   node fix.js\n\n"
        "The selected cells go in; what it prints replaces them.", initialvalue=settings.get("last_command", ""),
        parent=app)
    if not cmd:
        return None
    settings.set("last_command", cmd)
    ok, msg = run_command(app, cmd)
    (messagebox.showinfo if ok else messagebox.showwarning)("Run command", msg, parent=app)
    return ok
