"""Canvas spreadsheet grid: draws only visible cells; cell/range selection; in-place editor."""
import csv
import io
import re
import sys
import tkinter as tk
from tkinter import font as tkfont
from tkinter import ttk

from simplecore.theme import C, display_fill, near_black, text_color

import autocomplete
from cell_editor import CellEditor
from formatting import DEFAULT_SIZE, parse_number, render
from formula import quote_sheet, references, shift

ALL = 1 << 20          # bound used for whole-row / whole-column selections
REF_COLORS = ["#1a73e8", "#d93025", "#9334e6", "#188038", "#e8710a", "#007b83", "#b06000"]
EXTRA_ROWS, EXTRA_COLS = 50, 10   # scroll room past the data


def col_name(i):
    """0 -> A, 25 -> Z, 26 -> AA"""
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def blend(c1, c2, a):
    a1 = [int(c1[i:i + 2], 16) for i in (1, 3, 5)]
    a2 = [int(c2[i:i + 2], 16) for i in (1, 3, 5)]
    return "#%02x%02x%02x" % tuple(round(x + (y - x) * a) for x, y in zip(a1, a2))


class Grid(ttk.Frame):
    """app must provide: set_cells([(r, c, text), ...]) and on_select()."""

    def __init__(self, master, app, scale):
        super().__init__(master)
        self.app, self.scale = app, scale
        self.sheet = None
        self.top = self.left = 0
        self.cur = self.anchor = self.ext = (0, 0)   # active cell, selection corners
        self.extra = []                              # more blocks selected with Ctrl+click
        self.selected_chart = self.selected_image = None   # a picked chart / picture (arrows move it)
        self.ncols = 0
        self.merge_at = {}
        self.ref_marks = []   # [(r1, c1, r2, c2, color)] dashed boxes for the formula being typed
        self.point = None     # (entry, start, end, anchor, ref) of the reference last inserted by clicking
        self.handle = None    # (x, y) of the fill handle as last drawn
        self.border_lines = []
        self.row_drag = None  # (row, height px) while a row edge is dragged
        self.fill_preview = None   # (r1, c1, r2, c2) dashed box while dragging the fill handle
        self.rowpx, self.big, self.style_cols = {}, {}, 0   # row height cache, rows with big fonts
        self.wraps, self.wrap_cache = {}, {}   # row -> wrapped columns; wrap_lines() results
        self.editor = None
        self.drag = None
        self.view = (range(0), [])
        self.ui_font = tkfont.nametofont("TkDefaultFont")
        self.base_font = tkfont.Font(**self.ui_font.actual())   # headers; follows zoom
        self.fonts = {}
        self.zoom = 1.0
        self.rowh = int(self.base_font.metrics("linespace") * 1.5)

        self.cv = tk.Canvas(self, bg=C["cell_bg"], highlightthickness=0, takefocus=1)
        self.ys = ttk.Scrollbar(self, orient="vertical", command=self.yview)
        self.xs = ttk.Scrollbar(self, orient="horizontal", command=self.xview)
        self.cv.grid(row=0, column=0, sticky="nsew")
        self.ys.grid(row=0, column=1, sticky="ns")
        self.xs.grid(row=1, column=0, sticky="ew")
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        cv = self.cv
        cv.bind("<Configure>", lambda e: self.redraw())
        cv.bind("<Button-1>", self.on_press)
        cv.bind("<Shift-Button-1>", lambda e: self.on_press(e, extend=True))
        cv.bind("<Control-Button-1>", lambda e: self.on_press(e, add=True))   # one more block
        cv.bind("<B1-Motion>", self.on_drag)
        cv.bind("<ButtonRelease-1>", self.on_release)
        cv.bind("<Double-1>", self.on_double)
        cv.bind("<Motion>", self.on_motion)
        cv.bind("<Leave>", lambda e: self.notes.hide())
        from comments import HoverNote
        self.notes = HoverNote(self)
        cv.bind("<MouseWheel>", self.on_wheel)
        cv.bind("<Shift-MouseWheel>", lambda e: self.on_wheel(e, horizontal=True))
        cv.bind("<Control-MouseWheel>", lambda e: self.app.zoom_step(1 if e.delta > 0 else -1))
        import platform_ui  # Linux: the wheel comes as buttons 4 / 5
        platform_ui.bind_wheel(cv, self.on_wheel, lambda e: self.on_wheel(e, horizontal=True),
                               lambda e: self.app.zoom_step(1 if e.delta > 0 else -1))
        platform_ui.bind_touchpad(cv, self.on_touchpad)   # macOS two-finger scroll (Tk 9)
        for key, d in {"Up": (-1, 0), "Down": (1, 0), "Left": (0, -1), "Right": (0, 1)}.items():
            cv.bind(f"<{key}>", lambda e, d=d: self.move(*d))
            cv.bind(f"<Shift-{key}>", lambda e, d=d: self.move(*d, extend=True))
            cv.bind(f"<Control-{key}>", lambda e, d=d: self.move(*d, jump=True))
            cv.bind(f"<Control-Shift-{key}>", lambda e, d=d: self.move(*d, extend=True, jump=True))
        cv.bind("<Prior>", lambda e: self.move(-self.page(), 0))
        cv.bind("<Next>", lambda e: self.move(self.page(), 0))
        cv.bind("<Shift-Prior>", lambda e: self.move(-self.page(), 0, extend=True))
        cv.bind("<Shift-Next>", lambda e: self.move(self.page(), 0, extend=True))
        cv.bind("<Home>", lambda e: self.goto(self.cur[0], 0))
        cv.bind("<Control-Home>", lambda e: self.goto(0, 0))
        cv.bind("<Control-End>", lambda e: self.goto(max(len(self.sheet.rows) - 1, 0), max(self.ncols - 1, 0)))
        cv.bind("<Return>", lambda e: self.move(1, 0))
        cv.bind("<Shift-Return>", lambda e: self.move(-1, 0))
        cv.bind("<Tab>", lambda e: self.move(0, 1))
        cv.bind("<Shift-Tab>", lambda e: self.move(0, -1))
        cv.bind("<F2>", lambda e: self.edit())
        cv.bind("<Alt-Down>", lambda e: self.open_dropdown())
        cv.bind("<Escape>", lambda e: (self.app.cancel_painter(), self.unpick()))
        cv.bind("<Delete>", lambda e: self.clear())
        cv.bind("<BackSpace>", lambda e: self.backspace())
        cv.bind("<Control-a>", lambda e: self.smart_select_all())
        cv.bind("<Control-c>", lambda e: self.copy())
        cv.bind("<Control-x>", lambda e: self.cut())
        cv.bind("<Control-v>", lambda e: self.paste())
        cv.bind("<Control-V>", lambda e: self.paste("values"))   # Ctrl+Shift+V
        cv.bind("<Button-3>", self.on_right)
        cv.bind("<Key>", self.on_key)

    # ---------- data ----------

    def load(self, sheet):
        self.commit()
        self.selected_chart = self.selected_image = None
        if self.sheet is not None:
            self.sheet.view = (self.top, self.left, self.cur, self.anchor, self.ext)
        self.sheet = sheet
        self.top, self.left, self.cur, self.anchor, self.ext = sheet.view or (0, 0, (0, 0), (0, 0), (0, 0))
        self.reload()
        self.cv.focus_set()
        self.app.on_select()

    def reload(self):
        """Rebuild lookups after rows/columns moved (keeps scroll position)."""
        sheet = self.sheet
        self.merge_at, self.rowpx, self.big, self.wraps = {}, {}, {}, {}
        self.style_cols = 0
        for (r, c), st in sheet.styles.items():
            self.style_cols = max(self.style_cols, c + 1)
            if st and st.get("size", 0) > max(DEFAULT_SIZE, self.big.get(r, 0)):
                self.big[r] = st["size"]
            if st and st.get("wrap"):
                self.wraps.setdefault(r, set()).add(c)
        for m in sheet.merges:
            r1, c1, r2, c2 = m
            if (r2 - r1 + 1) * (c2 - c1 + 1) <= 20000:
                for r in range(r1, r2 + 1):
                    for c in range(c1, c2 + 1):
                        self.merge_at[(r, c)] = m
        self.data_changed()

    def data_changed(self):
        self.ncols = max((len(r) for r in self.sheet.rows), default=0)
        if self.wraps:            # wrapped text may need a different row height now
            self.rowpx = {}
        self.redraw()

    def text(self, r, c):
        """What's stored / edited (a formula stays '=...')."""
        rows = self.sheet.rows
        return rows[r][c] if r < len(rows) and c < len(rows[r]) else ""

    def value(self, r, c):
        """What's shown: formulas are calculated."""
        t = self.text(r, c)
        if len(t) > 1 and t[0] == "=" and not getattr(self.app, "show_formulas", False):
            return self.app.engine.display(self.sheet, r, c)
        if not t and getattr(self.app.engine, "spill", None):   # filled by a dynamic-array formula
            v = self.app.engine.spilled(self.sheet, r, c)
            if v is not None:
                return v
        return t

    def style(self, r, c):
        return self.sheet.styles.get((r, c))

    def span(self, r, c):
        return self.merge_at.get((r, c)) or (r, c, r, c)

    # ---------- geometry ----------

    def px(self, n):
        return int(n * self.scale * self.zoom)

    def zoomed(self, size):
        """Font size (points, or negative pixels) at the current zoom, never 0."""
        z = round(size * self.zoom)
        return z or (1 if size > 0 else -1)

    def set_zoom(self, zoom):
        self.commit()
        self.zoom = min(max(zoom, 0.25), 4.0)
        self.fonts, self.rowpx = {}, {}
        self.base_font.configure(size=self.zoomed(self.ui_font.actual()["size"]))
        self.rowh = int(self.base_font.metrics("linespace") * 1.5)
        self.redraw()

    def colw(self, c):
        s = self.sheet
        if c in s.hidden_cols:
            return 0
        if c in s.px:
            return int(s.px[c] * self.zoom)   # stored at 100%
        if c in s.widths:
            return self.px(s.widths[c] * 7 + 5)
        return self.px(72)

    def headings(self):
        return not (self.sheet and self.sheet.props.get("headings") is False)

    def hdr_h(self):
        """Height of the column-letter header (0 when headings are hidden)."""
        return self.rowh if self.headings() else 0

    def hdr_w(self):
        if not self.headings():
            return 0
        digits = len(str(self.top + self.vis_rows() + 1))
        return self.base_font.measure("0" * max(digits, 2)) + self.px(14)

    def rh(self, r):
        """Row height in px: grows for big fonts or a custom height from the file; 0 if hidden."""
        if r in self.sheet.hidden_rows or r in self.sheet.filtered_rows:
            return 0
        if self.row_drag and self.row_drag[0] == r:      # being resized right now
            return self.row_drag[1]
        if r in self.sheet.heights:                       # a set height is used as it is
            return max(int(self.sheet.heights[r] * 96 / 72 * self.scale * self.zoom), 2)
        h = self.rowpx.get(r)
        if h is None:
            h = self.rowh
            if r in self.big:
                h = max(h, int(self.font({"size": self.big[r]}).metrics("linespace") * 1.35))
            for c in self.wraps.get(r, ()):   # wrapped text: tall enough for all its lines
                text = self.value(r, c)
                if text and (r, c) not in self.merge_at and c not in self.sheet.hidden_cols:
                    font = self.font(self.style(r, c))
                    n = len(self.wrap_lines(text, font, self.colw(c) - 2 * self.px(4)))
                    h = max(h, min(n * font.metrics("linespace") + self.px(8), self.px(545)))
            self.rowpx[r] = h
        return h

    def wrap_lines(self, text, font, width):
        """Split text into lines that fit `width` px: at line breaks, then between words,
        and inside a word only when the word alone is too long."""
        key = (text, str(font), width)
        if key in self.wrap_cache:
            return self.wrap_cache[key]
        lines = []
        for para in text.split("\n"):
            line = ""
            for word in re.split(r"(\s+)", para):
                if not word:
                    continue
                if font.measure(line + word) <= width or not line:
                    line += word
                else:
                    lines.append(line.rstrip())
                    line = word.lstrip()
                while line and font.measure(line) > width and len(line) > 1:   # one long word
                    cut = len(line) - 1
                    while cut > 1 and font.measure(line[:cut]) > width:
                        cut -= 1
                    lines.append(line[:cut])
                    line = line[cut:]
            lines.append(line.rstrip())
        if len(self.wrap_cache) > 5000:
            self.wrap_cache.clear()
        self.wrap_cache[key] = lines
        return lines

    # Frozen panes: rows 0..fr-1 and columns 0..fc-1 always stay on screen; scrolling starts after
    # them. self.top / self.left are the first *scrolled* row / column (never inside the frozen part).

    @property
    def fr(self):
        return self.sheet.freeze[0] if self.sheet else 0

    @property
    def fc(self):
        return self.sheet.freeze[1] if self.sheet else 0

    def frozen_h(self):
        return sum(self.rh(i) for i in range(self.fr))

    def frozen_w(self):
        return sum(self.colw(i) for i in range(self.fc))

    def clamp_scroll(self):
        self.top, self.left = max(self.top, self.fr), max(self.left, self.fc)

    def row_seq(self):
        """Rows in screen order: frozen ones, then from self.top on."""
        yield from range(self.fr)
        r = max(self.top, self.fr)
        while True:
            yield r
            r += 1

    def col_seq(self):
        yield from range(self.fc)
        c = max(self.left, self.fc)
        while True:
            yield c
            c += 1

    def row_y(self, r):
        """y of row r's top edge (above the scrolled area when scrolled out of view)."""
        if r < self.fr:
            return self.hdr_h() + sum(self.rh(i) for i in range(r))
        base = self.hdr_h() + self.frozen_h()
        if r >= self.top:
            return base + sum(self.rh(i) for i in range(self.top, r))
        return base - sum(self.rh(i) for i in range(r, self.top))

    def vis_rows(self):
        """Scrollable rows that fit fully under the frozen ones."""
        H, y, n = self.cv.winfo_height(), self.hdr_h() + self.frozen_h(), 0
        while y + self.rh(self.top + n) <= H:
            y += self.rh(self.top + n)
            n += 1
        return max(n, 1)

    def style_changed(self, rows):
        """Recompute big-font rows after styles changed."""
        for (_r, c) in self.sheet.styles:
            self.style_cols = max(self.style_cols, c + 1)
        for r in set(rows):
            styles = {c: st for c in range(self.style_cols) if (st := self.sheet.styles.get((r, c)))}
            big = max((st.get("size", 0) for st in styles.values()), default=0)
            if big > DEFAULT_SIZE:
                self.big[r] = big
            else:
                self.big.pop(r, None)
            wraps = {c for c, st in styles.items() if st.get("wrap")}
            if wraps:
                self.wraps[r] = wraps
            else:
                self.wraps.pop(r, None)
            self.rowpx.pop(r, None)
        self.redraw()

    def page(self):
        return max(self.vis_rows() - 1, 1)

    def col_x(self, c):
        x = self.hdr_w()
        if c < self.fc:
            return x + sum(self.colw(i) for i in range(c))
        base = x + self.frozen_w()
        if c >= self.left:
            return base + sum(self.colw(i) for i in range(self.left, c))
        return base - sum(self.colw(i) for i in range(c, self.left))

    def hit(self, x, y):
        """-> (row, col); -1 means the header."""
        r = -1
        if y >= self.hdr_h():
            ry = self.hdr_h()
            for r in self.row_seq():
                if ry + self.rh(r) > y:
                    break
                ry += self.rh(r)
        if x < self.hdr_w():
            return r, -1
        cx = self.hdr_w()
        for c in self.col_seq():
            if cx + self.colw(c) > x:
                break
            cx += self.colw(c)
        return r, c

    def row_border_at(self, x, y):
        """Row whose bottom header edge is under the mouse, for resizing."""
        if x >= self.hdr_w() or y < self.hdr_h():
            return None
        for r in self.view[0]:
            h = self.rh(r)
            if h and abs(y - (self.row_y(r) + h)) <= self.px(3):
                return r
        return None

    def border_at(self, x, y):
        """Column whose right header edge is under the mouse, for resizing."""
        if y >= self.hdr_h():
            return None
        for c, cx, w in self.view[1]:
            if abs(x - (cx + w)) <= self.px(3):
                return c
        return None

    def see(self, r, c):
        """Scroll so cell (r, c) is visible (frozen rows / columns always are)."""
        r, c = min(r, ALL - 1), min(c, ALL - 1)
        fr, fc = self.fr, self.fc
        if r >= fr:
            if r < self.top:
                self.top = r
            elif r >= self.top + self.vis_rows():
                avail, t = self.cv.winfo_height() - self.hdr_h() - self.frozen_h(), r   # r = last full row
                total = self.rh(r)
                while t > fr and total + self.rh(t - 1) <= avail:
                    t -= 1
                    total += self.rh(t)
                self.top = t
        if c >= fc:
            if c < self.left:
                self.left = c
            else:
                avail = self.cv.winfo_width() - self.hdr_w() - self.frozen_w()
                while self.left < c and sum(self.colw(i) for i in range(self.left, c + 1)) > avail:
                    self.left += 1
        self.clamp_scroll()

    def yview(self, *a):
        total = max(len(self.sheet.rows), self.top + self.vis_rows()) + EXTRA_ROWS if self.sheet else 1
        if a[0] == "moveto":
            self.top = int(float(a[1]) * total)
        else:
            self.top += int(a[1]) * (self.page() if a[2] == "pages" else 1)
        self.clamp_scroll()
        self.redraw_soon()

    def xview(self, *a):
        total = max(self.ncols, self.left + len(self.view[1])) + EXTRA_COLS if self.sheet else 1
        if a[0] == "moveto":
            self.left = int(float(a[1]) * total)
        else:
            self.left += int(a[1]) * (max(len(self.view[1]) - 1, 1) if a[2] == "pages" else 1)
        self.clamp_scroll()
        self.redraw_soon()

    # ---------- selection ----------

    def selected_lines(self):
        """('col', c1, c2) when whole columns are selected, ('row', r1, r2) for whole rows, else None."""
        r1, c1, r2, c2 = self.sel()
        if r1 == 0 and r2 >= ALL and c2 < ALL:
            return "col", c1, c2
        if c1 == 0 and c2 >= ALL and r2 < ALL:
            return "row", r1, r2
        return None

    def on_selected_header(self, x, y):
        """Is (x, y) on the header of a selected whole column / row (which can be dragged to move it)?"""
        lines = self.selected_lines() if not self.extra else None
        if not lines or (x >= self.hdr_w() and y >= self.hdr_h()):
            return False
        r, c = self.hit(x, y)
        axis, i1, i2 = lines
        return (axis == "col" and r < 0 <= c and i1 <= c <= i2) or (axis == "row" and c < 0 <= r and i1 <= r <= i2)

    def select_line(self, axis, i1, i2):
        """Select whole columns (axis 'col') or rows i1 .. i2."""
        self.extra = []
        if axis == "col":
            self.cur, self.anchor, self.ext = (0, i1), (0, i1), (ALL, i2)
        else:
            self.cur, self.anchor, self.ext = (i1, 0), (i1, 0), (i2, ALL)
        self.redraw()
        self.app.on_select()

    def line_drop(self, axis, x, y):
        """Where dragged rows / columns would go: the line boundary nearest the mouse (index of the line
        they'd sit before)."""
        W, H = self.cv.winfo_width(), self.cv.winfo_height()
        r, c = self.hit(min(max(x, self.hdr_w() + 1), W - 1), min(max(y, self.hdr_h() + 1), H - 1))
        if axis == "col":
            if c < 0:
                return None
            return c + 1 if x > self.col_x(c) + self.colw(c) / 2 else c
        if r < 0:
            return None
        return r + 1 if y > self.row_y(r) + self.rh(r) / 2 else r

    def sel(self):
        (ar, ac), (er, ec) = self.anchor, self.ext
        return min(ar, er), min(ac, ec), max(ar, er), max(ac, ec)

    def multi(self):
        return self.anchor != self.ext or bool(self.extra)

    def rects(self):
        """Every selected block: the ones added with Ctrl+click, then the current one."""
        return self.extra + [self.sel()]

    def in_selection(self, r, c):
        return any(r1 <= r <= r2 and c1 <= c <= c2 for r1, c1, r2, c2 in self.rects())

    def cells_rects(self):
        """sel_cells() for every selected block (for formatting / clearing several blocks at once)."""
        return [self.sel_cells(rect) for rect in self.rects()]

    def sel_data(self):
        """Selection clipped to the data (rows may be empty if nothing there)."""
        r1, c1, r2, c2 = self.sel()
        return r1, c1, min(r2, len(self.sheet.rows) - 1), min(c2, self.ncols - 1)

    def sel_bounded(self):
        """Selection as-is, except whole rows/columns stop at the data."""
        r1, c1, r2, c2 = self.sel()
        if r2 >= ALL:
            r2 = max(len(self.sheet.rows) - 1, r1)
        if c2 >= ALL:
            c2 = max(self.ncols - 1, c1)
        return r1, c1, r2, c2

    def sel_cells(self, rect=None):
        """Selection for formatting: whole rows/columns stop at the data (or formatted cells);
        a normal block is used as selected, empty cells included."""
        r1, c1, r2, c2 = rect or self.sel()
        if r2 >= ALL:
            r2 = max(len(self.sheet.rows) - 1, r1)
        if c2 >= ALL:
            c2 = max(self.ncols - 1, self.style_cols - 1, c1)
        return r1, c1, r2, c2

    def goto(self, r, c):
        self.commit()
        self.extra = []
        self.cur = self.anchor = self.ext = (max(r, 0), max(c, 0))
        self.see(*self.cur)
        self.redraw_selection()
        self.app.on_select()
        return "break"

    def select(self, r1, c1, r2, c2):
        self.extra = []
        self.cur = self.anchor = (r1, c1)
        self.ext = (r2, c2)
        self.see(r1, c1)
        self.redraw()
        self.app.on_select()

    def select_all(self):
        self.commit()
        self.extra = []
        self.cur = self.anchor = (0, 0)
        self.ext = (ALL, ALL)
        self.redraw()
        self.app.on_select()
        return "break"

    def region(self, r, c):
        """Excel 'current region': the block of non-empty cells touching (r, c), diagonals included."""
        filled = lambda r, c: r >= 0 and c >= 0 and self.text(r, c) != ""
        r1 = r2 = r
        c1 = c2 = c
        grown = True
        while grown:
            grown = False
            while any(filled(r1 - 1, cc) for cc in range(c1 - 1, c2 + 2)):
                r1, grown = r1 - 1, True
            while any(filled(r2 + 1, cc) for cc in range(c1 - 1, c2 + 2)):
                r2, grown = r2 + 1, True
            while any(filled(rr, c1 - 1) for rr in range(r1 - 1, r2 + 2)):
                c1, grown = c1 - 1, True
            while any(filled(rr, c2 + 1) for rr in range(r1 - 1, r2 + 2)):
                c2, grown = c2 + 1, True
        return r1, c1, r2, c2

    def smart_select_all(self):
        """Ctrl+A like Excel: first the data block around the cursor, pressed again selects everything."""
        self.commit()
        r, c = self.cur
        reg = self.region(r, c)
        if reg == self.sel() or (reg[:2] == reg[2:] and not self.text(r, c)):
            return self.select_all()
        self.anchor, self.ext = reg[:2], reg[2:]
        self.redraw()
        self.app.on_select()
        return "break"

    def jump(self, r, c, dr, dc):
        """Ctrl+arrow: to the edge of the current data block."""
        maxr, maxc = max(len(self.sheet.rows) - 1, 0), max(self.ncols - 1, 0)

        def full(r, c):
            return 0 <= r <= maxr and 0 <= c <= maxc and self.text(r, c) != ""

        if full(r, c) and full(r + dr, c + dc):
            while full(r + dr, c + dc):
                r, c = r + dr, c + dc
            return r, c
        r, c = r + dr, c + dc
        while 0 <= r <= maxr and 0 <= c <= maxc and not full(r, c):
            r, c = r + dr, c + dc
        return min(max(r, 0), maxr), min(max(c, 0), maxc)

    def visible(self, r, c, dr=1, dc=1):
        """Nearest cell that isn't in a hidden row/column, going on in the move's direction
        (or back the other way at the sheet's start)."""
        hr, hc = self.sheet.hidden_rows | self.sheet.filtered_rows, self.sheet.hidden_cols
        for step in (dr or 1, -(dr or 1)):
            rr = r
            while rr in hr and rr >= 0:
                rr += step
            if rr >= 0:
                r = rr
                break
        for step in (dc or 1, -(dc or 1)):
            cc = c
            while cc in hc and cc >= 0:
                cc += step
            if cc >= 0:
                c = cc
                break
        return r, c

    def fix_cursor(self):
        """After hiding: move the cursor off hidden rows/columns (selection kept otherwise)."""
        if self.visible(*self.cur) != self.cur:
            self.cur = self.anchor = self.ext = self.visible(*self.cur)
            self.see(*self.cur)

    def move(self, dr, dc, extend=False, jump=False):
        im = getattr(self, "selected_image", None)
        if im is not None and im in (self.sheet.extra.get("images") or []) and not jump:
            import images
            images.nudge(self.app, im, dr, dc, resize=extend)
            return "break"
        ch = getattr(self, "selected_chart", None)
        if ch is not None and ch in self.sheet.charts and not jump:   # arrows move a picked chart
            import charts
            charts.nudge(self.app, ch, dr, dc, resize=extend)
            return "break"
        self.commit()
        if not self.sheet:
            return "break"
        r, c = self.ext if extend else self.cur
        r, c = min(r, len(self.sheet.rows) + EXTRA_ROWS), min(c, self.ncols + EXTRA_COLS)
        if jump:
            r, c = self.jump(r, c, dr, dc)
        else:
            if not extend:  # step out of a merged block
                r1, c1, r2, c2 = self.span(r, c)
                r = r2 if dr > 0 else r1 if dr < 0 else r
                c = c2 if dc > 0 else c1 if dc < 0 else c
            r, c = r + dr, c + dc
        r, c = self.visible(max(r, 0), max(c, 0), dr, dc)
        self.extra = []
        if extend:
            self.ext = (r, c)
        else:
            self.cur = self.anchor = self.ext = (r, c)
        self.see(r, c)
        self.redraw_selection()
        self.app.on_select()
        return "break"

    # ---------- drawing ----------

    def font(self, st):
        key = tuple(bool(st and st.get(k)) for k in ("bold", "italic", "underline", "strike")) \
            + ((st or {}).get("size", DEFAULT_SIZE), (st or {}).get("font"))
        if key not in self.fonts:
            spec = self.ui_font.actual()
            spec.update(weight="bold" if key[0] else "normal", slant="italic" if key[1] else "roman",
                        underline=key[2], overstrike=key[3], size=self.zoomed(spec["size"] * key[4] / DEFAULT_SIZE))
            if key[5] and key[5].lower() in self.installed_fonts():
                spec["family"] = key[5]
            self.fonts[key] = tkfont.Font(**spec)
        return self.fonts[key]

    def installed_fonts(self):
        if not hasattr(self, "_families"):
            self._families = {f.lower() for f in tkfont.families(self)}
        return self._families

    def fit(self, s, font, avail, is_num):
        if avail <= 0:
            return ""
        if font.measure(s) <= avail:
            return s
        if is_num:
            return "#" * max(avail // font.measure("#"), 1)
        lo, hi = 0, len(s)
        while lo < hi:
            mid = (lo + hi + 1) // 2
            if font.measure(s[:mid]) <= avail:
                lo = mid
            else:
                hi = mid - 1
        return s[:lo]

    def redraw(self):
        cv = self.cv
        if getattr(self, "_redraw_job", None) is not None:      # drawn now: the queued redraw isn't needed
            self._root().after_cancel(self._redraw_job)
            self._redraw_job = None
        cv.delete("all")
        if not self.sheet:
            return
        W, H, rh = cv.winfo_width(), cv.winfo_height(), self.hdr_h()
        hw = self.hdr_w()
        self.clamp_scroll()
        ys, y = {}, rh                   # row -> (y, height), in screen order (frozen rows first)
        for r in self.row_seq():
            if y >= H:
                break
            ys[r] = (y, self.rh(r))
            y += ys[r][1]
        rows = list(ys)
        limit = getattr(self.sheet, "loaded_limit", None)
        if limit and rows and len(self.sheet.rows) >= limit and max(rows) >= len(self.sheet.rows) - 30:
            self._root().after_idle(self.app.load_more_rows)     # scrolled near the end of a long sheet: read more
        cols, x = [], hw
        for c in self.col_seq():
            if x >= W:
                break
            w = self.colw(c)
            cols.append((c, x, w))
            x += w
        self.view = (rows, cols)
        if not cols:
            return
        xs = {c: (x, w) for c, x, w in cols}
        freeze_y, freeze_x = rh + self.frozen_h(), hw + self.frozen_w()
        self.border_lines = []   # filled by draw_cell, drawn after every cell (so fills can't cover them)
        self.script_cells = self.app_script_cells()
        if self.sheet.props.get("grid") is not False:   # gridlines: one long line per row / column, not a box per cell
            gl = C["grid_line"]
            right, bottom = min(x, W), min(y, H)
            for r in rows:
                yy = ys[r][0] + ys[r][1]
                cv.create_line(hw, yy, right, yy, fill=gl)
            for _c, cx, w in cols:
                cv.create_line(cx + w, rh, cx + w, bottom, fill=gl)
        for r in rows:
            y, h = ys[r]
            for c, x, w in cols:
                if (r, c) not in self.merge_at:
                    self.draw_cell(r, c, x, y, w, h, cols)
        for r1, c1, r2, c2 in self.sheet.merges:
            if any(r1 <= r <= r2 for r in ys) and any(c1 <= c <= c2 for c in xs):
                x1, x2 = self.col_x(c1), self.col_x(c2 + 1)
                y1, y2 = self.row_y(r1), self.row_y(r2 + 1)
                if r1 >= self.fr:   # partly scrolled away: don't cover the frozen rows / columns
                    y1 = max(y1, freeze_y)
                if c1 >= self.fc:
                    x1 = max(x1, freeze_x)
                if x2 > x1 and y2 > y1:
                    self.draw_cell(r1, c1, x1, y1, x2 - x1, y2 - y1, None)
        for args in self.border_lines:
            self.draw_border(*args)
        cv.tag_raise("txt")
        import charts
        import images
        images.draw_all(self)                  # pictures and charts float over the cells
        charts.draw_all(self, C)
        self.filter_buttons = {}   # column -> (x1, y1, x2, y2) of its ▼ button
        f = self.sheet.filter
        if f and f["range"][0] in ys:
            fr1, fc1, _, fc2 = f["range"]
            y, h = ys[fr1]
            size = max(min(h - 4, self.px(16)), 8)
            for c, x, w in cols:
                if fc1 <= c <= fc2 and w > size:
                    x2, y2 = x + w - 2, y + h - 2
                    btn = (x2 - size, y2 - size, x2, y2)
                    on = c in f["criteria"]
                    cv.create_rectangle(*btn, fill=C["accent"] if on else C["header_bg"],
                                        outline=C["header_line"], tags="filterbtn")
                    cv.create_text((btn[0] + btn[2]) / 2, (btn[1] + btn[3]) / 2, text="▼",
                                   fill=C["cell_bg"] if on else C["header_fg"],
                                   font=(self.base_font.actual("family"), max(int(size * 0.45), 6)),
                                   tags="filterbtn")
                    self.filter_buttons[c] = btn
        self.table_buttons = []    # (box, table): ▼ of Excel tables that aren't the live filter
        if self.sheet.extra.get("tables"):
            import tables
            for t, tr, (tc1, tc2) in tables.header_buttons(self.sheet):
                if tr not in ys:
                    continue
                y, h = ys[tr]
                size = max(min(h - 4, self.px(16)), 8)
                for c, x, w in cols:
                    if tc1 <= c <= tc2 and w > size:
                        x2, y2 = x + w - 2, y + h - 2
                        btn = (x2 - size, y2 - size, x2, y2)
                        cv.create_rectangle(*btn, fill=C["header_bg"], outline=C["header_line"], tags="filterbtn")
                        cv.create_text((btn[0] + btn[2]) / 2, (btn[1] + btn[3]) / 2, text="▼", fill=C["header_fg"],
                                       font=(self.base_font.actual("family"), max(int(size * 0.45), 6)),
                                       tags="filterbtn")
                        self.table_buttons.append((btn, t, c))
        if self.fr:   # freeze lines
            cv.create_line(0, freeze_y, W, freeze_y, fill=C["muted"], width=2, tags="freeze")
        if self.fc:
            cv.create_line(freeze_x, 0, freeze_x, H, fill=C["muted"], width=2, tags="freeze")
        if getattr(self.app, "page_breaks", False) and self.app.current is self.sheet:
            import print_preview
            print_preview.draw_breaks(self, cv, ys, cols, W, H)

        self._geom = (rows, cols, ys, xs, W, H, rh, hw)
        self._layout = self.layout_key()
        self._was_multi = self.multi() or bool(self.ref_marks) or bool(self.fill_preview)
        self.draw_overlay(*self._geom)
        self.keep_editor_in_place()

        total_r = max(len(self.sheet.rows), self.top + len(rows), 1) + EXTRA_ROWS
        total_c = max(self.ncols, self.left + len(cols), 1) + EXTRA_COLS
        self.ys.set(self.top / total_r, (self.top + len(rows)) / total_r)
        self.xs.set(self.left / total_c, (self.left + len(cols)) / total_c)

    def app_script_cells(self):
        """Cells of this sheet with a click script (cell_scripts.py)."""
        app = getattr(self, "app", None)
        if app is None or not getattr(app, "book_props", None) or not app.book_props.get("scripts"):
            return set()
        import cell_scripts
        return cell_scripts.cells(app, self.sheet)

    def layout_key(self):
        """What the cells on screen depend on; while it's unchanged only the overlay needs drawing."""
        return (id(self.sheet), self.top, self.left, self.zoom, self.cv.winfo_width(), self.cv.winfo_height(),
                self.fr, self.fc, bool(self.editor))

    def redraw_selection(self):
        """Cursor moved without scrolling: redraw only the selection, headers and handles (fast), else all."""
        g = getattr(self, "_geom", None)
        if (g is None or self.layout_key() != getattr(self, "_layout", None) or self.multi() or self._was_multi
                or self.ref_marks or self.fill_preview or not self.sheet):
            return self.redraw()
        self.cv.delete("ov")
        self.draw_overlay(*g)

    def draw_overlay(self, rows, cols, ys, xs, W, H, rh, hw):
        """Everything that follows the selection: outlines, active cell, ▼ of a list cell, fill handle,
        highlighted headers, group buttons. Tagged "ov" so redraw_selection can replace just these."""
        cv = self.cv
        first_item = cv.create_line(0, 0, 0, 0, fill="")

        def box(r1, c1, r2, c2, width, color=None, dash=None):
            vr = [r for r in rows if r1 <= r <= r2]
            vc = [c for c, _, _ in cols if c1 <= c <= c2]
            if not vr or not vc:
                return
            x1 = xs[vc[0]][0] - (0 if vc[0] == c1 else 3)    # edges running off screen stick out
            x2 = sum(xs[vc[-1]]) + (0 if vc[-1] == c2 else 3)
            y1 = ys[vr[0]][0] - (0 if vr[0] == r1 else 3)
            y2 = sum(ys[vr[-1]]) + (0 if vr[-1] == r2 else 3)
            cv.create_rectangle(x1, y1, x2, y2, outline=color or C["accent"], width=width, dash=dash,
                                tags="ref" if dash else "")

        sr1, sc1, sr2, sc2 = self.sel()
        for rect in self.extra:
            box(*rect, 1)
        if self.multi():
            box(sr1, sc1, sr2, sc2, 1)
        box(*self.span(*self.cur), 2)
        self.dv_button = None   # drop-down arrow of a "list" data validation, on the active cell
        ar1, ac1, ar2, ac2 = self.span(*self.cur)
        dv = self.app.rules.validation_at(self.sheet, ar1, ac1)
        if dv and dv["type"] == "list" and dv["arrow"] and not self.editor and ar2 in ys and ac2 in xs:
            x2, y2 = sum(xs[ac2]), sum(ys[ar2])
            size = max(min(ys[ar2][1] - 4, self.px(16)), 8)
            btn = (x2 + 2, y2 - size - 1, x2 + 2 + size, y2 - 1)    # just right of the cell, like Excel
            cv.create_rectangle(*btn, fill=C["header_bg"], outline=C["header_line"], tags="dvbtn")
            cv.create_text((btn[0] + btn[2]) / 2, (btn[1] + btn[3]) / 2, text="▼", fill=C["header_fg"],
                           font=(self.base_font.actual("family"), max(int(size * 0.45), 6)), tags="dvbtn")
            self.dv_button = btn
        for r1, c1, r2, c2, color in self.ref_marks:   # cells used by the formula being typed
            box(r1, c1, r2, c2, 2, color, (4, 3))
        if self.fill_preview:
            box(*self.fill_preview, 2, "#5f6368", (4, 3))
        # fill handle: small square at the selection's bottom-right corner
        hr1, hc1, hr2, hc2 = (sr1, sc1, sr2, sc2) if self.multi() else self.span(*self.cur)
        self.handle = None
        if hr2 < ALL and hc2 < ALL and hr2 in ys and hc2 in xs and not self.editor and not self.extra:
            hx, hy, s = sum(xs[hc2]), sum(ys[hr2]), max(self.px(3), 3)
            cv.create_rectangle(hx - s, hy - s, hx + s, hy + s, fill=C["accent"], outline=C["cell_bg"], tags="handle")
            self.handle = (hx, hy)

        # headers (none when the sheet hides them)
        heads = self.headings()
        if heads:
            cv.create_rectangle(0, 0, W, rh, fill=C["header_bg"], outline="")
            cv.create_rectangle(0, 0, hw, H, fill=C["header_bg"], outline="")
        hidden_c, hidden_r, mark = self.sheet.hidden_cols, self.sheet.hidden_rows, max(self.px(2), 2)
        for c, x, w in (cols if heads else []):
            if w == 0:
                continue
            on = any(b <= c <= d for _, b, _, d in self.rects())
            cv.create_rectangle(x, 0, x + w, rh, fill=C["header_sel"] if on else C["header_bg"], outline=C["header_line"])
            cv.create_text(x + w // 2, rh // 2, text=col_name(c), fill=C["header_fg"], font=self.base_font)
            if c - 1 in hidden_c:   # marker: hidden column(s) just left of here
                cv.create_rectangle(x - mark, 2, x + mark, rh - 2, fill=C["hidden_mark"], outline="", tags="hidden")
        for r in (rows if heads else []):
            y, h = ys[r]
            if h == 0:
                continue
            on = any(a <= r <= b for a, _, b, _ in self.rects())
            cv.create_rectangle(0, y, hw, y + h, fill=C["header_sel"] if on else C["header_bg"], outline=C["header_line"])
            cv.create_text(hw - self.px(6), y + h // 2, text=str(r + 1), anchor="e",
                           fill=C["header_fg"], font=self.base_font)
            if r - 1 in hidden_r:   # marker: hidden row(s) just above here
                cv.create_rectangle(2, y - mark, hw - 2, y + mark, fill=C["hidden_mark"], outline="", tags="hidden")
        if heads:
            cv.create_rectangle(0, 0, hw, rh, fill=C["header_bg"], outline=C["header_line"])
        import outline
        outline.draw(self, ys, xs, C)          # group −/+ buttons
        for i in cv.find_all():
            if i >= first_item:
                cv.addtag_withtag("ov", i)

    def draw_border(self, side, line, x, y, w, h):
        """One side of a cell border: width by line style, dashes, double = two thin lines."""
        style, color = line
        if C["dark"] and near_black(color):   # black lines would vanish on dark cells
            color = C["cell_fg"]
        width = 3 if style == "thick" else 2 if style.startswith("medium") else 1
        dash = (2, 2) if style in ("dotted", "hair") else (5, 3) if "dash" in style.lower() else None
        x1, y1, x2, y2 = {"top": (x, y, x + w, y), "bottom": (x, y + h, x + w, y + h),
                          "left": (x, y, x, y + h), "right": (x + w, y, x + w, y + h)}[side]
        if style == "double":
            dx, dy = (0, 1) if side in ("top", "bottom") else (1, 0)
            for k in (-1, 1):
                self.cv.create_line(x1 + k * dx, y1 + k * dy, x2 + k * dx, y2 + k * dy, fill=color, tags="border")
            return
        self.cv.create_line(x1, y1, x2, y2, fill=color, width=width, dash=dash, tags="border")

    def draw_cell(self, r, c, x, y, w, h, cols):
        cv = self.cv
        if w == 0 or h == 0:   # hidden row / column
            return
        st = self.style(r, c)
        if self.sheet.extra.get("tables"):                   # Excel table banding under the cell's own look
            import tables
            band = tables.look(self.sheet, r, c)
            if band:
                st = {**band, **(st or {})}
        cf = self.app.rules.conditional(self.sheet, r, c)   # conditional formatting from the file
        bar = None
        if cf:
            bar = cf.pop("bar", None)
            st = {**(st or {}), **cf}
        real_bg = st.get("bg") if st else None
        bg = display_fill(real_bg)                          # dark mode: light fills shown as a dark shade
        if self.multi() and self.in_selection(r, c) and self.span(*self.cur)[:2] != (r, c):
            bg = blend(bg or C["cell_bg"], C["accent"], 0.15)
        if cols is None and not bg:        # a merged block hides the gridlines inside it
            bg = C["cell_bg"]
        if bg:                             # plain cells: just the gridlines drawn by redraw
            grid = C["grid_line"] if self.sheet.props.get("grid") is not False else bg
            cv.create_rectangle(x, y, x + w, y + h, fill=bg, outline=grid)
        if bar:   # data bar: a coloured bar as long as the value is big
            frac, color = bar
            cv.create_rectangle(x + 2, y + 3, x + 2 + max(int((w - 4) * frac), 1), y + h - 3,
                                fill=blend(color, C["cell_bg"], 0.45), outline=color, tags="databar")
        if (r, c) in self.script_cells:     # accent corner = clicking runs its script
            k = max(self.px(7), 5)
            cv.create_polygon(x + 1, y + 1, x + k, y + 1, x + 1, y + k, fill=C["accent"], outline="", tags="script")
        if (r, c) in self.sheet.comments:   # red corner = has a comment
            k = max(self.px(6), 5)
            cv.create_polygon(x + w - k, y + 1, x + w - 1, y + 1, x + w - 1, y + k, fill="#d93025",
                              outline="", tags="note")
        for side, line in ((st or {}).get("border") or {}).items():
            self.border_lines.append((side, line, x, y, w, h))
        text = self.value(r, c)
        if hasattr(text, "url"):                    # =IMAGE(): the picture in the cell (alt text until it's there)
            import cell_images
            if cell_images.draw(self, text, x, y, w, h):
                return
            st_, _ = cell_images.state(text.url)
            text = text or ("Loading picture…" if st_ == "loading" else "#picture?")
        if not text:
            return
        disp = render(text, st.get("numfmt")) if st else text
        is_num = parse_number(text) is not None
        align = (st.get("align") if st else None) or ("right" if is_num else "left")
        font = self.font(st)
        pad = self.px(4)
        avail = w - 2 * pad
        fg = text_color(st.get("fg") if st else None, real_bg)
        just = "center" if align == "center" else "right" if align == "right" else "left"
        tx = x + w // 2 if just == "center" else x + w - pad if just == "right" else x + pad
        indent = self.px(9) * int(st.get("indent", 0)) if st else 0   # Excel indent: about one character each
        if indent:
            tx = tx + indent if just == "left" else tx - indent if just == "right" else tx
            avail -= indent
        valign = (st.get("valign") if st else None) or "bottom"   # Excel's default
        vpad = self.px(3)
        ty = y + vpad if valign == "top" else y + h // 2 if valign == "center" else y + h - vpad
        # Tk anchor = vertical part (n / - / s) + horizontal part (w / - / e)
        anchor = ({"top": "n", "center": "", "bottom": "s"}[valign]
                  + {"left": "w", "center": "", "right": "e"}[just]) or "center"
        if self.sheet.rich and not (st and st.get("wrap")):   # several fonts in the cell: run by run
            import richtext
            runs = richtext.valid(self.sheet, r, c, self.text(r, c))
            if runs and not getattr(self.app, "show_formulas", False):
                return richtext.draw(self, runs, st, x, y, w, h, fg, valign) and None
        if st and st.get("wrap") and not is_num:   # wrapped: several lines inside the cell
            lines = self.wrap_lines(disp, font, avail)
            fits = max((h - self.px(4)) // font.metrics("linespace"), 1)
            cv.create_text(tx, ty, text="\n".join(lines[:fits]), anchor=anchor, justify=just,
                           font=font, fill=fg, tags="txt")
            return
        disp = disp.split("\n", 1)[0]
        if align == "left" and cols:  # spill into empty cells on the right, like Excel
            for c2_, _, w2 in cols:
                if c2_ <= c:
                    continue
                if self.text(r, c2_) or (r, c2_) in self.merge_at:
                    break
                avail += w2
        disp = self.fit(disp, font, avail, is_num)
        angle = st.get("rotate") if st else None
        if angle:                                    # rotated text: drawn turned around its middle
            cv.create_text(x + w // 2, y + h // 2, text=disp, angle=angle, font=font, fill=fg, tags="txt")
            return
        cv.create_text(tx, ty, text=disp, anchor=anchor, font=font, fill=fg, tags="txt")

    # ---------- mouse ----------

    def on_press(self, e, extend=False, add=False):
        pe = self.point_entry()
        if pe is not None:   # typing a formula: the click inserts a reference instead
            r, c = self.hit(e.x, e.y)
            if (r, c) != (-1, -1):
                self.point_at(pe, r, c, extend)
                self.drag = ("point", pe)
            return "break"
        if self.app.formula_home is not None:   # a click that isn't pointing finishes the formula
            self.app.value_commit()
            return "break"
        for c, (x1, y1, x2, y2) in getattr(self, "filter_buttons", {}).items():   # a filter ▼ button
            if x1 <= e.x <= x2 and y1 <= e.y <= y2:
                self.commit()
                self.app.open_filter(c, e.x_root, e.y_root)
                return "break"
        for (x1, y1, x2, y2), t, c in getattr(self, "table_buttons", []):   # an Excel table's ▼
            if x1 <= e.x <= x2 and y1 <= e.y <= y2:
                self.commit()
                import tables
                tables.start_filter(self.app, t)
                self.app.open_filter(c, e.x_root, e.y_root)
                return "break"
        import charts
        import images
        im = images.hit(self, e.x, e.y)
        if im is not None and not extend:        # a picture: pick it; drag moves it, its corner resizes it
            self.commit()
            box = next(b for b, i in self.image_items if i is im)
            corner = images.on_corner(box, e.x, e.y)
            self.selected_image, self.selected_chart = im, None
            self.drag = ("image", im, "resize" if corner else "move", e.x, e.y)
            self.cv.focus_set()
            self.redraw()
            return "break"
        if getattr(self, "selected_image", None) is not None:
            self.selected_image = None
        ch = charts.hit(self, e.x, e.y)
        if ch is not None and not extend:        # a chart: pick it; drag moves it, its corner resizes it
            self.commit()
            box = next(b for b, c in self.chart_boxes if c is ch)
            corner = images.on_corner(box, e.x, e.y)
            self.selected_chart = ch
            self.drag = ("chart", ch, "resize" if corner else "move", e.x, e.y)
            self.cv.focus_set()
            self.redraw()
            return "break"
        if getattr(self, "selected_chart", None) is not None:
            self.selected_chart = None
        import outline
        if outline.click(self, e.x, e.y):        # a group's −/+ button
            self.commit()
            return "break"
        b = getattr(self, "dv_button", None)
        if b and b[0] <= e.x <= b[2] and b[1] <= e.y <= b[3]:   # data validation ▼: pick from the list
            self.commit()
            self.open_dropdown()
            return "break"
        if self.on_handle(e.x, e.y):            # start dragging the fill handle
            r1, c1, r2, c2 = self.sel() if self.multi() else self.span(*self.cur)
            self.drag = ("fill", (r1, c1, r2, c2))
            return "break"
        if self.on_sel_edge(e.x, e.y):          # grab the selection by its edge: drag to move it
            self.commit()
            rect = self.sel() if self.multi() else self.span(*self.cur)
            self.drag = ("move", rect, self.hit(e.x, e.y))
            at = self.hit(e.x, e.y)               # (not moved in the end: a click, which runs a cell's script)
            self.script_press = self.span(*at)[:2] if min(at) >= 0 and not (extend or add) and \
                self.span(*at)[:2] in getattr(self, "script_cells", ()) else None
            return "break"
        self.commit()
        self.cv.focus_set()
        if not self.sheet:
            return "break"
        b = self.border_at(e.x, e.y)
        if b is not None:
            self.drag = ("resize", b, e.x, self.colw(b))
            return "break"
        rb = self.row_border_at(e.x, e.y)
        if rb is not None:
            self.drag = ("rowresize", rb, e.y, self.rh(rb))
            return "break"
        r, c = self.hit(e.x, e.y)
        if r < 0 and c < 0:
            return self.select_all()
        grab = None if extend or add or self.extra else self.selected_lines()
        if grab and ((grab[0] == "col" and r < 0 and grab[1] <= c <= grab[2]) or
                     (grab[0] == "row" and c < 0 and grab[1] <= r <= grab[2])):
            # pressed on a header that is already selected: dragging moves those columns / rows
            self.drag = ("movelines", grab[0], grab[1], grab[2], c if grab[0] == "col" else r)
            self.line_target = None
            return "break"
        if add:          # Ctrl+click: keep what's selected, start another block
            if self.sel() not in self.extra:
                self.extra.append(self.sel())
        elif not extend:
            self.extra = []
        if r < 0:      # column header
            self.drag = "cols"
            if extend:
                self.ext = (ALL, c)
            else:
                self.cur, self.anchor, self.ext = (0, c), (0, c), (ALL, c)
        elif c < 0:    # row header
            self.drag = "rows"
            if extend:
                self.ext = (r, ALL)
            else:
                self.cur, self.anchor, self.ext = (r, 0), (r, 0), (r, ALL)
        else:
            self.drag = "cells"
            if extend:
                self.ext = (r, c)
            else:
                self.cur = self.anchor = self.ext = (r, c)
        # a plain click on a cell with a script runs it when the button comes up on the same cell
        self.script_press = self.span(r, c)[:2] if r >= 0 and c >= 0 and not (extend or add) and \
            self.span(r, c)[:2] in getattr(self, "script_cells", ()) else None
        self.redraw()
        self.app.on_select()
        return "break"

    def unpick(self):
        """Esc: let go of a picked chart / picture."""
        if getattr(self, "selected_chart", None) is not None or getattr(self, "selected_image", None) is not None:
            self.selected_chart = self.selected_image = None
            self.redraw()

    def on_sel_edge(self, x, y):
        """True on the outline of the selection (not the fill handle), where a drag moves the cells."""
        if not self.sheet or self.extra:
            return False
        r1, c1, r2, c2 = self.sel() if self.multi() else self.span(*self.cur)
        if r2 >= ALL or c2 >= ALL:
            return False
        x1, x2, y1, y2 = self.col_x(c1), self.col_x(c2 + 1), self.row_y(r1), self.row_y(r2 + 1)
        k = 3
        near_x = abs(x - x1) <= k or abs(x - x2) <= k
        near_y = abs(y - y1) <= k or abs(y - y2) <= k
        inside_x, inside_y = x1 - k <= x <= x2 + k, y1 - k <= y <= y2 + k
        return (near_x and inside_y) or (near_y and inside_x)

    def move_target(self, x, y, drag):
        """Top-left cell the dragged block would land on."""
        _, (r1, c1, r2, c2), (gr, gc) = drag
        W, H = self.cv.winfo_width(), self.cv.winfo_height()
        r, c = self.hit(min(max(x, self.hdr_w() + 1), W - 1), min(max(y, self.hdr_h() + 1), H - 1))
        if r < 0 or c < 0:
            return None
        return max(r1 + r - gr, 0), max(c1 + c - gc, 0)

    def on_handle(self, x, y):
        return self.handle is not None and abs(x - self.handle[0]) <= self.px(5) \
            and abs(y - self.handle[1]) <= self.px(5)

    def fill_target(self, x, y):
        """(direction, count) for the fill handle dragged to x, y; None if still inside the block."""
        r1, c1, r2, c2 = self.drag[1]
        W, H = self.cv.winfo_width(), self.cv.winfo_height()
        r, c = self.hit(min(max(x, self.hdr_w()), W - 1), min(max(y, self.hdr_h()), H - 1))
        down, up, right, left = r - r2, r1 - r, c - c2, c1 - c
        best = max(down, up, right, left)
        if best <= 0:
            return None
        return {down: "down", up: "up", right: "right", left: "left"}[best], best

    def width_chars(self, px):
        """Column width in pixels at 100% zoom -> Excel's unit (characters), as saved in files."""
        return max(round((px / self.scale - 5) / 7, 2), 0.5)

    def on_release(self, e):
        drag, self.drag = self.drag, None
        pressed, self.script_press = getattr(self, "script_press", None), None
        at = self.hit(e.x, e.y) if pressed else (-1, -1)
        clicked = drag == "cells" or (isinstance(drag, tuple) and drag[0] == "move" and drag[2] == at)
        if pressed and clicked and not self.multi() and min(at) >= 0 and self.span(*at)[:2] == pressed:
            import cell_scripts  # a click (not a drag) on a script cell: run its script
            self._root().after_idle(lambda: cell_scripts.run(self.app, *pressed))
        if isinstance(drag, tuple) and drag[0] == "movelines":
            _, axis, a, b, start = drag
            t, self.line_target = getattr(self, "line_target", None), None
            if t is not None and not a <= t <= b + 1:
                self.app.move_lines(axis, a, b - a + 1, t)
            else:                                  # a click, not a drag: select just that line, as before
                self.select_line(axis, start, start)
            return
        if isinstance(drag, tuple) and drag[0] == "chart":
            import charts
            charts.drop(self, drag[1], drag[2], drag[3], drag[4], e.x, e.y)
            return
        if isinstance(drag, tuple) and drag[0] == "image":
            import images
            images.drop(self, drag[1], drag[2], drag[3], drag[4], e.x, e.y)
            return
        if drag in ("cells", "rows", "cols") and getattr(self.app, "painter", None):
            self.app.paint_format()                  # format painter: this selection gets the format
            return
        if isinstance(drag, tuple) and drag[0] == "move":   # dragged the selection's edge: move the cells
            self.fill_preview = None
            target = self.move_target(e.x, e.y, drag)
            if target is not None and target != drag[1][:2]:
                self.app.move_block(drag[1], *target)
            else:
                self.redraw()
            return
        if isinstance(drag, tuple) and drag[0] == "rowresize":   # dragged height: a real change, in points
            rd, self.row_drag = self.row_drag, None
            if rd is not None:
                self.app.set_height(rd[0], round(rd[1] / (self.scale * self.zoom) * 72 / 96, 2))
            return
        if isinstance(drag, tuple) and drag[0] == "resize":   # dragged width becomes a real change
            c = drag[1]
            px = self.sheet.px.pop(c, None)
            if px is not None:
                self.app.set_width(c, self.width_chars(px))
            return
        if isinstance(drag, tuple) and drag[0] == "fill":
            self.drag = drag
            target = self.fill_target(e.x, e.y)
            self.drag, self.fill_preview = None, None
            if target:
                self.app.fill(drag[1], *target, copy=bool(e.state & 0x4))   # Ctrl: copy, don't count
            else:
                self.redraw()

    def on_drag(self, e):
        if not self.drag:
            return
        if isinstance(self.drag, tuple) and self.drag[0] == "move":   # preview where the block lands
            t = self.move_target(e.x, e.y, self.drag)
            r1, c1, r2, c2 = self.drag[1]
            self.fill_preview = (t[0], t[1], t[0] + r2 - r1, t[1] + c2 - c1) if t else None
            self.redraw()
            return
        if isinstance(self.drag, tuple) and self.drag[0] == "fill":   # preview the fill area
            r1, c1, r2, c2 = self.drag[1]
            t = self.fill_target(e.x, e.y)
            if t is None:
                self.fill_preview = None
            else:
                d, n = t
                self.fill_preview = {"down": (r2 + 1, c1, r2 + n, c2), "up": (max(r1 - n, 0), c1, r1 - 1, c2),
                                     "right": (r1, c2 + 1, r2, c2 + n), "left": (r1, max(c1 - n, 0), r2, c1 - 1)}[d]
            self.redraw()
            return
        if isinstance(self.drag, tuple) and self.drag[0] == "rowresize":
            _, r, y0, h0 = self.drag
            self.row_drag = (r, max(h0 + e.y - y0, self.px(4)))
            self.redraw()
            return
        if isinstance(self.drag, tuple) and self.drag[0] == "point":   # drag out a range reference
            W, H = self.cv.winfo_width(), self.cv.winfo_height()
            r, c = self.hit(min(max(e.x, self.hdr_w()), W - 1), min(max(e.y, self.hdr_h()), H - 1))
            self.point_at(self.drag[1], r, c, extend=True)
            return
        if isinstance(self.drag, tuple) and self.drag[0] == "movelines":   # line where the rows / columns land
            _, axis, a, b, _start = self.drag
            W, H = self.cv.winfo_width(), self.cv.winfo_height()
            if axis == "col":                      # past an edge: scroll that way
                if e.x >= W:
                    self.left += 1
                elif e.x < self.hdr_w() + self.frozen_w() and self.left > self.fc:
                    self.left -= 1
            else:
                if e.y >= H:
                    self.top += 1
                elif e.y < self.hdr_h() + self.frozen_h() and self.top > self.fr:
                    self.top -= 1
            self.line_target = self.line_drop(axis, e.x, e.y)
            self.redraw()
            t = self.line_target
            if t is not None and not a <= t <= b + 1:
                if axis == "col":
                    x = self.col_x(t)
                    self.cv.create_line(x, 0, x, H, fill=C["accent"], width=3, tags="moveline")
                else:
                    y = self.row_y(t)
                    self.cv.create_line(0, y, W, y, fill=C["accent"], width=3, tags="moveline")
            return
        if isinstance(self.drag, tuple) and self.drag[0] in ("chart", "image"):   # outline where it lands
            kind, obj, mode, x0, y0 = self.drag
            boxes = self.chart_boxes if kind == "chart" else getattr(self, "image_items", [])
            box = next((b for b, o in boxes if o is obj), None)
            self.cv.delete("dragpreview")
            if box is not None:
                dx, dy = e.x - x0, e.y - y0
                x1, y1, x2, y2 = box
                if mode == "move":
                    rect = (x1 + dx, y1 + dy, x2 + dx, y2 + dy)
                elif kind == "image":                    # pictures keep their proportions
                    import images
                    k = images.resize_factor(self, obj, dx, dy)
                    rect = (x1, y1, x1 + (x2 - x1) * k, y1 + (y2 - y1) * k)
                else:
                    rect = (x1, y1, max(x2 + dx, x1 + 8), max(y2 + dy, y1 + 8))
                self.cv.create_rectangle(*rect, outline=C["accent"], dash=(4, 2), width=2, tags="dragpreview")
            return
        if isinstance(self.drag, tuple) and self.drag[0] == "resize":  # column resize
            _, c, x0, w0 = self.drag
            self.sheet.px[c] = max(w0 + e.x - x0, self.px(8)) / self.zoom
            self.rowpx = {}    # wrapped text re-flows
            self.redraw()
            return
        W, H, hw, rh = self.cv.winfo_width(), self.cv.winfo_height(), self.hdr_w(), self.hdr_h()
        if self.drag != "cols":   # auto-scroll when dragging past an edge
            if e.y >= H:
                self.top += 1
            elif e.y < rh + self.frozen_h() and self.top > self.fr:
                self.top -= 1
        if self.drag != "rows":
            if e.x >= W:
                self.left += 1
            elif e.x < hw + self.frozen_w() and self.left > self.fc:
                self.left -= 1
        r, c = self.hit(min(max(e.x, hw), W - 1), min(max(e.y, rh), H - 1))
        self.ext = (ALL, c) if self.drag == "cols" else (r, ALL) if self.drag == "rows" else (r, c)
        self.redraw()
        self.app.on_select()

    def on_double(self, e):
        if self.editor is not None or self.point:   # double-click while pointing a reference
            return "break"
        if self.on_handle(e.x, e.y):                 # fill down to the length of the next column
            self.drag = None
            self.app.fill_to_neighbour()
            return "break"
        rb = self.row_border_at(e.x, e.y)
        if rb is not None:                           # row edge: back to automatic height
            self.drag = self.row_drag = None
            self.app.set_height(rb, None)
            return "break"
        b = self.border_at(e.x, e.y)
        if b is not None:  # autofit column
            fonts = [self.font(self.style(r, b)) for r in range(min(len(self.sheet.rows), 2000))]
            widths = [f.measure(render(self.value(r, b),(self.style(r, b) or {}).get("numfmt")))
                      for r, f in enumerate(fonts)]
            self.app.set_width(b, self.width_chars((max(widths + [self.px(30)]) + self.px(10)) / self.zoom))
            return "break"
        if e.y >= self.hdr_h() and e.x >= self.hdr_w():
            return self.edit()

    def on_right(self, e):
        """Right-click: keep the selection if clicked inside it, else select there; then the menu."""
        self.commit()
        self.cv.focus_set()
        if not self.sheet:
            return "break"
        import charts
        ch = charts.hit(self, e.x, e.y)
        if ch is not None:                        # right-click on a chart: its own small menu
            m = tk.Menu(self, tearoff=False)
            m.add_command(label="Edit chart…", command=lambda: charts.EditDialog(self.app, ch))
            m.add_command(label="Delete chart", command=lambda: charts.delete(self.app, ch))
            self.chart_menu = m
            try:
                m.tk_popup(e.x_root, e.y_root)
            finally:
                m.grab_release()
            return "break"
        r, c = self.hit(e.x, e.y)
        r1, c1, r2, c2 = self.sel()
        inside =(r < 0 or r1 <= r <= r2) and (c < 0 or c1 <= c <= c2)
        full_col, full_row = r2 >= ALL, c2 >= ALL
        if r < 0 and c >= 0:
            inside = full_col and c1 <= c <= c2
        elif c < 0 and r >= 0:
            inside = full_row and r1 <= r <= r2
        if not inside:
            self.on_press(e)
            self.drag = None
        self.app.context_menu(e.x_root, e.y_root)
        return "break"

    def on_motion(self, e):
        self.notes.motion(e.x, e.y)
        import images
        over = images.cursor(self, e.x, e.y)       # a picture / chart: move or resize pointer
        if over:
            self.cv["cursor"] = over
        elif self.on_selected_header(e.x, e.y) and self.border_at(e.x, e.y) is None \
                and self.row_border_at(e.x, e.y) is None:
            self.cv["cursor"] = "fleur"                # selected column / row header: drag to move it
        elif self.on_handle(e.x, e.y):
            self.cv["cursor"] = "crosshair"
        elif self.row_border_at(e.x, e.y) is not None:
            self.cv["cursor"] = "sb_v_double_arrow"
        elif getattr(self, "script_cells", None) and self.border_at(e.x, e.y) is None and \
                min(self.hit(e.x, e.y)) >= 0 and self.span(*self.hit(e.x, e.y))[:2] in self.script_cells:
            self.cv["cursor"] = "hand2"                # a cell that runs a script when clicked
        else:
            self.cv["cursor"] = "sb_h_double_arrow" if self.border_at(e.x, e.y) is not None else ""

    def typing_formula(self):
        """A formula is being typed in the cell (scrolling must keep it open to point at far cells)."""
        e = self.editor
        return e is not None and e.winfo_exists() and e.get().startswith("=")

    def keep_editor_in_place(self):
        """After scrolling: the in-cell editor follows its cell; when the cell is out of view it waits out of
        sight (still focused, still typing; the formula bar shows the text) until the cell scrolls back."""
        e = self.editor
        if e is None or not e.winfo_exists():
            return
        sheet, r, c = e.cell
        if sheet is not self.sheet:
            return
        x, y = self.col_x(c), self.row_y(r)
        top = self.hdr_h() + (self.frozen_h() if r >= self.fr else 0)
        left = self.hdr_w() + (self.frozen_w() if c >= self.fc else 0)
        if y < top or x < left or y >= self.cv.winfo_height() or x >= self.cv.winfo_width():
            x, y = -10000, -10000
        info = e.place_info()
        if (str(info.get("x")), str(info.get("y"))) != (str(x), str(y)):
            e.place_configure(x=x, y=y)

    def on_wheel(self, e, horizontal=False):
        if not self.typing_formula():          # typing "=…": scroll to point at cells further away
            self.commit()
        # a notch (delta 120) = 3 rows; touchpads send small deltas, which add up instead of jumping 3 rows each
        import platform_ui
        old_mac = platform_ui.old_mac_wheel()                              # Tk 8.6 on macOS: delta 1, 2, 3…
        rows = -e.delta if old_mac else -e.delta * 3 / 120
        acc = rows / 3 if horizontal else rows
        if abs(e.delta) >= 120 or old_mac:                                 # a notch: always moves, whole rows
            acc = round(acc) or (1 if acc > 0 else -1)
        else:
            acc += getattr(self, "_wheel_acc", 0.0)
        step = int(acc)
        self._wheel_acc = acc - step
        if not step:
            return "break"
        if horizontal:
            self.left += step
        else:
            self.top += step
        self.clamp_scroll()
        self.redraw_soon()
        return "break"

    def on_touchpad(self, e):
        """macOS two-finger scroll (Tk 9 <TouchpadScroll>): pixel deltas both ways, added up into rows / columns."""
        import platform_ui
        dx, dy = platform_ui.touchpad_deltas(int(e.delta))
        if not (dx or dy):
            return "break"
        if not self.typing_formula():
            self.commit()
        ax, ay = getattr(self, "_pad_acc", (0.0, 0.0))
        ax -= dx / max(self.colw(self.left) or self.px(64), 1)
        ay -= dy / max(self.rowh, 1)
        cols, rows = int(ax), int(ay)
        self._pad_acc = (ax - cols, ay - rows)
        if cols or rows:
            self.left += cols
            self.top += rows
            self.clamp_scroll()
            self.redraw_soon()
        return "break"

    def redraw_soon(self):
        """Redraw once the waiting events are handled: a fast wheel / scrollbar drag gives one redraw, not one
        per event (the redraws queued up and the sheet lagged behind the mouse)."""
        if getattr(self, "_redraw_job", None) is None:
            self._redraw_job = self._root().after_idle(self._redraw_now)

    def _redraw_now(self):
        self._redraw_job = None
        if self.cv.winfo_exists():
            self.redraw()

    # ---------- keyboard actions ----------

    def on_key(self, e):
        ctrl_only = e.state & 0x4 and not e.state & 0x20000   # Ctrl but not AltGr
        command = sys.platform == "darwin" and e.state & 0x8   # ⌘ + a key with no shortcut: don't type it
        if self.sheet and e.char and e.char.isprintable() and not ctrl_only and not command:
            return self.edit(e.char)

    def backspace(self, platform=None):
        """Backspace clears the active cell and starts editing. A Mac has no separate Delete key (its "delete"
        is Backspace): there a selected range, chart or picture is cleared / removed like Delete does."""
        if (platform or sys.platform) == "darwin" and (
                self.multi() or getattr(self, "selected_image", None) is not None
                or getattr(self, "selected_chart", None) is not None):
            return self.clear()
        return self.edit("")

    def clear(self):
        im = getattr(self, "selected_image", None)
        if im is not None and im in (self.sheet.extra.get("images") or []):   # Delete: the picked picture
            import images
            self.selected_image = None
            images.delete(self.app, im)
            return "break"
        ch = getattr(self, "selected_chart", None)
        if ch is not None and ch in self.sheet.charts:      # Delete with a chart picked removes the chart
            import charts
            self.selected_chart = None
            charts.delete(self.app, ch)
            return "break"
        changes = {}
        for rect in self.rects():
            r1, c1, r2, c2 = rect
            r2, c2 = min(r2, len(self.sheet.rows) - 1), min(c2, self.ncols - 1)
            changes.update({(r, c): (r, c, "") for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)
                            if self.text(r, c)})
        changes = list(changes.values())
        if changes:
            self.app.set_cells(changes)
        return "break"

    def copy(self):
        r1, c1, r2, c2 = self.sel_data()
        if not self.multi():
            r2, c2 = r1, c1   # single (maybe empty) cell
        buf = io.StringIO()   # other apps get the shown values
        w = csv.writer(buf, delimiter="\t", lineterminator="\r\n")
        for r in range(r1, r2 + 1):
            w.writerow([self.value(r, c) for c in range(c1, c2 + 1)])
        data = buf.getvalue().removesuffix("\r\n")
        self.clipboard_clear()
        self.clipboard_append(data)
        # pasting back here uses the formulas, with references moved like Excel
        raw = [[self.text(r, c) for c in range(c1, c2 + 1)] for r in range(r1, r2 + 1)]
        styles = [[self.sheet.styles.get((r, c)) for c in range(c1, c2 + 1)] for r in range(r1, r2 + 1)]
        self.clip = (data, r1, c1, raw, False, styles)
        return "break"

    def cut(self):
        self.copy()
        self.clip = self.clip[:4] + (True,) + self.clip[5:]   # cut + paste keeps references unchanged
        return self.clear()

    def paste(self, mode="all", transpose=False, op=None, skip_blanks=False):
        """mode: 'all', 'values' (shown values, no formats) or 'format' (formats only).
        transpose: rows become columns. op: 'add' / 'subtract' / 'multiply' / 'divide' the copied numbers
        into the cells pasted on. skip_blanks: empty copied cells leave the target cell alone."""
        self.commit()
        try:
            data = self.clipboard_get()
        except tk.TclError:
            return "break"
        clip = getattr(self, "clip", None)
        own = clip is not None and clip[0] == data
        if mode == "format" and not own:
            return "break"   # formats only exist for copies made here
        if own and mode != "values" and not op:   # our own copy: formulas + formats
            _, sr, sc, rows, cut, styles = clip
        else:
            rows = list(csv.reader(io.StringIO(data.rstrip("\r\n")), delimiter="\t")) or [[""]]
            sr = sc = cut = styles = None
        src_of = lambda i, j: (i, j)   # noqa: E731   (pasted position -> position in the copy)
        if transpose:
            w = max(map(len, rows))
            rows = [[rows[i][j] if j < len(rows[i]) else "" for i in range(len(rows))] for j in range(w)]
            if styles is not None:
                styles = [[styles[i][j] if j < len(styles[i]) else None for i in range(len(styles))]
                          for j in range(w)]
            src_of = lambda i, j: (j, i)   # noqa: E731
        if op or skip_blanks:
            return self.paste_combine(rows, op, skip_blanks, styles if mode != "values" else None)
        r0, c0, r2, c2 = self.sel()
        if len(rows) == 1 and len(rows[0]) == 1 and self.multi():  # one value fills the selection
            r2, c2 = min(r2, max(len(self.sheet.rows) - 1, r0)), min(c2, max(self.ncols - 1, c0))
        else:
            r2, c2 = r0 + len(rows) - 1, c0 + max(map(len, rows)) - 1
        changes = []
        for r in range(r0, r2 + 1):
            for c in range(c0, c2 + 1):
                i, j = (r - r0) % len(rows), (c - c0) % max(map(len, rows))
                if j >= len(rows[i]) or self.span(r, c)[:2] != (r, c):
                    continue
                v = rows[i][j]
                if sr is not None and not cut and v.startswith("="):
                    si, sj = src_of(i, j)
                    if transpose:
                        from formula import transpose_refs
                        v = transpose_refs(v, (sr + si, sc + sj), (r, c))
                    else:
                        v = shift(v, r - (sr + si), c - (sc + sj))
                if mode != "format":
                    changes.append(("v", r, c, v))
                if styles is not None:
                    changes.append(("s", r, c, styles[i][j]))
        if changes:
            self.app.apply_mixed(changes)
        self.select(r0, c0, r2, c2)
        return "break"

    def paste_combine(self, rows, op, skip_blanks, styles=None):
        """Paste special with an operation and / or skipping blanks (values only)."""
        r0, c0 = self.sel()[:2]
        changes = []
        for i, row in enumerate(rows):
            for j, v in enumerate(row):
                r, c = r0 + i, c0 + j
                if self.span(r, c)[:2] != (r, c) or (skip_blanks and v == ""):
                    continue
                if op:
                    p = parse_number(v)
                    if p is None:
                        continue
                    t = self.text(r, c)
                    cur = parse_number(self.value(r, c)) if t else 0.0
                    if cur is None:
                        continue
                    if op == "divide" and p == 0:
                        new = "#DIV/0!"
                    else:
                        x = {"add": cur + p, "subtract": cur - p, "multiply": cur * p,
                             "divide": cur / p if p else 0}[op]
                        new = str(int(x)) if float(x).is_integer() and abs(x) < 1e15 else f"{x:.15g}"
                    if t.startswith("=") and len(t) > 1 and new != "#DIV/0!":   # a formula keeps its formula
                        sym = {"add": "+", "subtract": "-", "multiply": "*", "divide": "/"}[op]
                        new = f"=({t[1:]}){sym}{v}"
                    changes.append(("v", r, c, new))
                else:
                    changes.append(("v", r, c, v))
                    if styles is not None and i < len(styles) and j < len(styles[i]):
                        changes.append(("s", r, c, styles[i][j]))
        if changes:
            self.app.apply_mixed(changes)
        self.select(r0, c0, r0 + len(rows) - 1, c0 + max(map(len, rows)) - 1)
        return "break"

    # ---------- in-place editor ----------

    def point_entry(self):
        """Entry (cell editor or formula bar) where a click should insert a reference, else None.
        Like Excel: only while typing a formula with the caret right after = ( , or an operator,
        or right after a reference that was just pointed at (a new click replaces it)."""
        e = self.editor
        if e is None and self.sheet and self.focus_get() is self.app.value:
            e = self.app.value
        if e is None or not e.get().startswith("="):
            return None
        caret = e.index("insert")
        if self.point and self.point[0] is e:
            _, start, end, _, ref = self.point
            if caret == end and e.get()[start:end] == ref:
                return e
        before = e.get()[:caret].rstrip()
        return e if before and before[-1] in "=(,;+-*/^&<>:%" else None

    def point_at(self, entry, r, c, extend=False):
        """Insert (or replace the last pointed) reference to cell / row (c=-1) / column (r=-1)."""
        last = self.point if self.point and self.point[0] is entry else None
        if last and (entry.index("insert") != last[2] or entry.get()[last[1]:last[2]] != last[4]):
            last = None
        if r >= 0 and c >= 0:
            r, c = self.span(r, c)[:2]
        anchor = last[3] if last and extend else (r, c)
        (ar, ac), (br, bc) = anchor, (r, c)
        if ar < 0 or br < 0:      # whole columns
            a, b = sorted((ac, bc))
            ref = f"{col_name(a)}:{col_name(b)}"
        elif ac < 0 or bc < 0:    # whole rows
            a, b = sorted((ar, br))
            ref = f"{a + 1}:{b + 1}"
        else:
            r1, r2, c1, c2 = min(ar, br), max(ar, br), min(ac, bc), max(ac, bc)
            ref = f"{col_name(c1)}{r1 + 1}" + ("" if (r1, c1) == (r2, c2) else f":{col_name(c2)}{r2 + 1}")
        home = self.app.home_sheet_name()
        if home and home != self.app.sheet:   # pointing at another sheet: Sheet!A1
            ref = f"{quote_sheet(self.app.sheet)}!{ref}"
        start = last[1] if last else entry.index("insert")
        end = last[2] if last else start
        entry.delete(start, end)
        entry.insert(start, ref)
        entry.icursor(start + len(ref))
        self.point = (entry, start, start + len(ref), anchor, ref)
        self.point_end = (r, c)
        entry.focus_set()
        if hasattr(entry, "fit"):
            entry.fit()
        self.show_refs(entry.get())

    def column_complete(self, e, ev=None):
        """Typing text that starts like a value already in this column: the rest is suggested (selected),
        Enter keeps it, typing on replaces it - like Excel's AutoComplete."""
        if ev is not None and (not ev.char or not ev.char.isprintable()):
            return
        text = e.get()
        if not text or text.startswith("=") or "\n" in text or parse_number(text) is not None:
            return
        if e.index("insert") != len(text):
            return
        _, r0, c = e.cell
        words = getattr(e, "column_words", None)
        if words is None:
            words, seen = [], set()
            rows = self.sheet.rows
            for r in range(max(0, r0 - 2000), min(len(rows), r0 + 2000)):
                t = rows[r][c] if r != r0 and c < len(rows[r]) else ""
                if t and not t.startswith("=") and parse_number(t) is None and t.casefold() not in seen:
                    seen.add(t.casefold())
                    words.append(t)
            e.column_words = words
        low = text.casefold()
        hits = [w for w in words if w.casefold().startswith(low) and len(w) > len(text)]
        if len({h.casefold() for h in hits}) != 1:
            return
        rest = hits[0][len(text):]
        e.insert("end", rest)
        e.icursor(len(text))
        e.select_range(len(text), len(text) + len(rest))
        if hasattr(e, "fit"):
            e.fit()

    def point_key(self, entry, dr, dc, extend=False):
        """Arrow key while typing a formula: insert / move a reference ("break"), or None if not pointing."""
        if self.point_entry() is not entry:
            return None
        last = self.point if self.point and self.point[0] is entry else None
        live = last and entry.index("insert") == last[2] and entry.get()[last[1]:last[2]] == last[4]
        if live and getattr(self, "point_end", None):
            base = self.point_end                 # keep moving the reference just pointed at
        else:
            base = entry.cell[1:] if hasattr(entry, "cell") else self.cur   # start at the edited cell
        r, c = max(base[0] + dr, 0), max(base[1] + dc, 0)
        self.point_at(entry, r, c, extend)
        self.see(r, c)
        return "break"

    def show_refs(self, text):
        """Dashed colored boxes around cells a formula uses (only refs to this sheet)."""
        marks, colors = [], {}
        here = (self.app.sheet or "").lower()
        home = (self.app.home_sheet_name() or "").lower()   # refs without "Sheet!" mean this sheet
        for sheet, r1, c1, r2, c2 in references(text):
            if (sheet.lower() if sheet is not None else home) != here:
                continue
            key = (r1, c1, r2, c2)
            if key not in colors:
                colors[key] = REF_COLORS[len(colors) % len(REF_COLORS)]
            marks.append(key + (colors[key],))
        if marks != self.ref_marks:
            self.ref_marks = marks
            self.redraw()

    def open_dropdown(self):
        """Alt+↓ or the ▼ beside a cell with a drop-down list (data validation): pick a value."""
        if not self.sheet:
            return "break"
        r1, c1, r2, _ = self.span(*self.cur)
        opts = self.app.rules.dropdown(self.sheet, r1, c1)
        if not opts:
            return "break"
        self.close_dropdown()
        top = tk.Toplevel(self)
        top.overrideredirect(True)
        top.attributes("-topmost", True)
        lb = tk.Listbox(top, height=min(len(opts), 10), activestyle="none", exportselection=False,
                        relief="solid", bd=1, highlightthickness=0, font=self.base_font,
                        width=max(len(o) for o in opts) + 2, background=C["field_bg"], foreground=C["fg"],
                        selectbackground=C["menu_active"], selectforeground=C["fg"])
        for o in opts:
            lb.insert("end", o)
        now = self.text(r1, c1)
        i = next((k for k, o in enumerate(opts) if o.lower() == now.lower()), 0)
        lb.selection_set(i)
        lb.see(i)
        lb.pack(fill="both", expand=True)
        top.update_idletasks()
        width = max(top.winfo_reqwidth(), self.col_x(self.span(*self.cur)[3] + 1) - self.col_x(c1))   # >= the cell
        top.geometry(f"{width}x{top.winfo_reqheight()}"
                     f"+{self.cv.winfo_rootx() + self.col_x(c1)}+{self.cv.winfo_rooty() + self.row_y(r2 + 1)}")

        def pick(_=None):
            sel = lb.curselection()
            self.close_dropdown()
            if sel and opts[sel[0]] != self.text(r1, c1):
                self.app.set_cells([(r1, c1, opts[sel[0]])])
            self.cv.focus_set()
            return "break"
        lb.bind("<ButtonRelease-1>", pick)
        lb.bind("<Return>", pick)
        lb.bind("<Escape>", lambda e: (self.close_dropdown(), self.cv.focus_set(), "break")[-1])
        lb.bind("<FocusOut>", lambda e: self._root().after(150, self.close_dropdown_if_unfocused))
        lb.focus_force()
        self.dropdown = (top, lb, opts, pick)
        return "break"

    def close_dropdown_if_unfocused(self):
        d = getattr(self, "dropdown", None)
        if d and self.focus_get() is not d[1]:
            self.close_dropdown()

    def close_dropdown(self):
        d, self.dropdown = getattr(self, "dropdown", None), None
        if d:
            d[0].destroy()

    def edit(self, initial=None):
        self.commit()
        if not self.sheet:
            return "break"
        r1, c1, r2, c2 = self.span(*self.cur)
        if self.sheet.props.get("protect") and not (self.style(r1, c1) or {}).get("unlocked"):
            self.app.protected_warning()          # protected sheet: locked cells can't be edited
            return "break"
        self.see(r1, c1)
        self.redraw()
        st = self.style(r1, c1)
        font = self.font(st)
        text = self.text(r1, c1) if initial is None else initial
        x, y = self.col_x(c1), self.row_y(r1)
        w = sum(self.colw(c) for c in range(c1, c2 + 1))
        h = self.row_y(r2 + 1) - y
        wrap = bool(st and st.get("wrap"))
        e = CellEditor(self.cv, font, wrap)
        e.cell = (self.sheet, r1, c1)
        e.insert(0, text)
        e.place(x=x, y=y)
        line_h = font.metrics("linespace")

        def fit(_=None):
            """Grow with the text: wider up to the window edge (or wrap at the cell width), taller per line."""
            room = max(self.cv.winfo_width() - x, w)
            if wrap:
                width = w
            else:
                longest = max((font.measure(t) for t in e.get().split("\n")), default=0)
                width = max(w, min(longest + self.px(20), room))
            e.place_configure(width=width)
            e.update_idletasks()
            e.place_configure(height=max(h, min(e.lines() * line_h + self.px(6), self.cv.winfo_height() - y)))
        fit()
        e.fit = fit
        e.focus_set()
        e.icursor("end")
        e.bind("<Return>", lambda ev: self.commit(1, 0))
        e.bind("<KP_Enter>", lambda ev: self.commit(1, 0))
        for seq in ("<Control-Return>", "<Alt-Return>"):   # a new line inside the cell
            e.bind(seq, lambda ev: (e.insert("insert", "\n"), fit(), "break")[-1])
        e.bind("<Shift-Return>", lambda ev: self.commit(-1, 0))
        e.bind("<Tab>", lambda ev: self.commit(0, 1))
        e.bind("<Shift-Tab>", lambda ev: self.commit(0, -1))
        e.bind("<Escape>", lambda ev: self.commit(cancel=True))
        e.bind("<FocusOut>", lambda ev: self._root().after_idle(self.focus_left, e))
        e.bind("<KeyRelease>", lambda ev: (fit(), self.show_refs(e.get())), add="+")
        e.bind("<KeyRelease>", lambda ev: self.column_complete(e, ev), add="+")
        typed = initial is not None
        for key, d in {"Up": (-1, 0), "Down": (1, 0), "Left": (0, -1), "Right": (0, 1)}.items():
            # in a formula, after = ( , or an operator: arrows point at cells (Shift: a range), like Excel;
            # otherwise typing over a cell: arrows finish the edit
            e.bind(f"<{key}>", lambda ev, d=d: self.point_key(e, *d) or (self.commit(*d) if typed else None))
            e.bind(f"<Shift-{key}>", lambda ev, d=d: self.point_key(e, *d, extend=True))
        import richtext
        runs = richtext.valid(self.sheet, r1, c1, text) if initial is None else None
        if runs:                              # several fonts in the cell: keep them while editing
            richtext.load_editor(self, e, runs, st)
        for k, key in (("b", "bold"), ("i", "italic"), ("u", "underline")):
            e.bind(f"<Control-{k}>", lambda ev, key=key: richtext.toggle(self, e, key) or "break")
        autocomplete.attach(e, self.app)    # function list / argument hints while typing a formula
        self.editor = e
        self.show_refs(text)
        return "break"

    def focus_left(self, e):
        """Editor lost focus: finish only if focus went to another control here - switching to
        another program (focus_get() is None) keeps the edit open, like Excel."""
        now = self.focus_get()
        if self.typing_formula() and now in (self.ys, self.xs):   # scrollbar while typing "=…": keep typing
            e.focus_set()
            return
        if self.editor is e and now not in (None, e):
            self.commit()

    def commit(self, dr=0, dc=0, cancel=False):
        e, self.editor = self.editor, None
        if e is None:
            return "break"
        sheet, r, c = e.cell
        val = e.get()
        import richtext
        rich = None
        if hasattr(e, "rich_base") or getattr(e, "rich_touched", False):   # the editor had / got rich text
            runs = richtext.runs_of_editor(e)
            rich = ("rt", r, c, None if richtext.plain(runs) or val.startswith("=") else runs)
        e.destroy()
        self.ref_marks, self.point, self.point_end = [], None, None
        rich_changed = rich is not None and rich[3] != sheet.rich.get((r, c))
        if not cancel and sheet is self.sheet and (val != self.text(r, c) or rich_changed):
            if val != self.text(r, c) and not self.app.allowed(r, c, val):   # data validation said no
                self.ref_marks = []
                self.goto(r, c)
                self.edit(val)
                return "break"
            st = self.style(r, c) or {}
            extra = [rich] if rich_changed else []
            if "\n" in val and not st.get("wrap"):   # a typed line break turns on wrap, like Excel
                self.app.apply_mixed([("v", r, c, val), ("s", r, c, {**st, "wrap": True})] + extra)
            elif extra:
                self.app.apply_mixed([("v", r, c, val)] + extra)
            else:
                self.app.set_cells([(r, c, val)])
        self.cv.focus_set()
        if dr or dc:
            self.move(dr, dc)
        else:
            self.redraw()
        return "break"
