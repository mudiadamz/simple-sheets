"""Canvas spreadsheet grid: draws only visible cells; cell/range selection; in-place editor."""
import csv
import io
import re
import tkinter as tk
from tkinter import font as tkfont
from tkinter import ttk

from formatting import DEFAULT_SIZE, parse_number, render
from cell_editor import CellEditor
from formula import quote_sheet, references, shift
from theme import C, text_color

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
        self.ncols = 0
        self.merge_at = {}
        self.ref_marks = []   # [(r1, c1, r2, c2, color)] dashed boxes for the formula being typed
        self.point = None     # (entry, start, end, anchor, ref) of the reference last inserted by clicking
        self.handle = None    # (x, y) of the fill handle as last drawn
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
        cv.bind("<B1-Motion>", self.on_drag)
        cv.bind("<ButtonRelease-1>", self.on_release)
        cv.bind("<Double-1>", self.on_double)
        cv.bind("<Motion>", self.on_motion)
        cv.bind("<MouseWheel>", self.on_wheel)
        cv.bind("<Shift-MouseWheel>", lambda e: self.on_wheel(e, horizontal=True))
        cv.bind("<Control-MouseWheel>", lambda e: self.app.zoom_step(1 if e.delta > 0 else -1))
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
        cv.bind("<Delete>", lambda e: self.clear())
        cv.bind("<BackSpace>", lambda e: self.edit(""))
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
        if len(t) > 1 and t[0] == "=":
            return self.app.engine.display(self.sheet, r, c)
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

    def hdr_w(self):
        digits = len(str(self.top + self.vis_rows() + 1))
        return self.base_font.measure("0" * max(digits, 2)) + self.px(14)

    def rh(self, r):
        """Row height in px: grows for big fonts or a custom height from the file; 0 if hidden."""
        if r in self.sheet.hidden_rows:
            return 0
        h = self.rowpx.get(r)
        if h is None:
            h = self.rowh
            if r in self.sheet.heights:
                h = max(h, int(self.sheet.heights[r] * 96 / 72 * self.scale * self.zoom))
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

    def row_y(self, r):
        """y of row r's top edge (negative when above the view)."""
        if r >= self.top:
            return self.rowh + sum(self.rh(i) for i in range(self.top, r))
        return self.rowh - sum(self.rh(i) for i in range(r, self.top))

    def vis_rows(self):
        H, y, n = self.cv.winfo_height(), self.rowh, 0
        while y + self.rh(self.top + n) <= H:
            y += self.rh(self.top + n)
            n += 1
        return max(n, 1)

    def style_changed(self, rows):
        """Recompute big-font rows after styles changed."""
        for (r, c) in self.sheet.styles:
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
        if c >= self.left:
            return x + sum(self.colw(i) for i in range(self.left, c))
        return x - sum(self.colw(i) for i in range(c, self.left))

    def hit(self, x, y):
        """-> (row, col); -1 means the header."""
        r = -1
        if y >= self.rowh:
            r, ry = self.top, self.rowh
            while ry + self.rh(r) <= y:
                ry += self.rh(r)
                r += 1
        if x < self.hdr_w():
            return r, -1
        c, cx = self.left, self.hdr_w()
        while cx + self.colw(c) <= x:
            cx += self.colw(c)
            c += 1
        return r, c

    def border_at(self, x, y):
        """Column whose right header edge is under the mouse, for resizing."""
        if y >= self.rowh:
            return None
        for c, cx, w in self.view[1]:
            if abs(x - (cx + w)) <= self.px(3):
                return c
        return None

    def see(self, r, c):
        r, c = min(r, ALL - 1), min(c, ALL - 1)
        if r < self.top:
            self.top = r
        elif r >= self.top + self.vis_rows():
            avail, t = self.cv.winfo_height() - self.rowh, r   # make r the last full row
            total = self.rh(r)
            while t > 0 and total + self.rh(t - 1) <= avail:
                t -= 1
                total += self.rh(t)
            self.top = t
        if c < self.left:
            self.left = c
        else:
            avail = self.cv.winfo_width() - self.hdr_w()
            while self.left < c and sum(self.colw(i) for i in range(self.left, c + 1)) > avail:
                self.left += 1

    def yview(self, *a):
        total = max(len(self.sheet.rows), self.top + self.vis_rows()) + EXTRA_ROWS if self.sheet else 1
        if a[0] == "moveto":
            self.top = int(float(a[1]) * total)
        else:
            self.top += int(a[1]) * (self.page() if a[2] == "pages" else 1)
        self.top = max(self.top, 0)
        self.redraw()

    def xview(self, *a):
        total = max(self.ncols, self.left + len(self.view[1])) + EXTRA_COLS if self.sheet else 1
        if a[0] == "moveto":
            self.left = int(float(a[1]) * total)
        else:
            self.left += int(a[1]) * (max(len(self.view[1]) - 1, 1) if a[2] == "pages" else 1)
        self.left = max(self.left, 0)
        self.redraw()

    # ---------- selection ----------

    def sel(self):
        (ar, ac), (er, ec) = self.anchor, self.ext
        return min(ar, er), min(ac, ec), max(ar, er), max(ac, ec)

    def multi(self):
        return self.anchor != self.ext

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

    def sel_cells(self):
        """Selection for formatting: whole rows/columns stop at the data, a lone empty cell still counts."""
        r1, c1, r2, c2 = self.sel()
        return (r1, c1, min(r2, max(len(self.sheet.rows) - 1, r1)),
                min(c2, max(self.ncols - 1, self.style_cols - 1, c1)))

    def goto(self, r, c):
        self.commit()
        self.cur = self.anchor = self.ext = (max(r, 0), max(c, 0))
        self.see(*self.cur)
        self.redraw()
        self.app.on_select()
        return "break"

    def select(self, r1, c1, r2, c2):
        self.cur = self.anchor = (r1, c1)
        self.ext = (r2, c2)
        self.see(r1, c1)
        self.redraw()
        self.app.on_select()

    def select_all(self):
        self.commit()
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
        hr, hc = self.sheet.hidden_rows, self.sheet.hidden_cols
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
        if extend:
            self.ext = (r, c)
        else:
            self.cur = self.anchor = self.ext = (r, c)
        self.see(r, c)
        self.redraw()
        self.app.on_select()
        return "break"

    # ---------- drawing ----------

    def font(self, st):
        key = tuple(bool(st and st.get(k)) for k in ("bold", "italic", "underline", "strike")) \
            + ((st or {}).get("size", DEFAULT_SIZE),)
        if key not in self.fonts:
            spec = self.ui_font.actual()
            spec.update(weight="bold" if key[0] else "normal", slant="italic" if key[1] else "roman",
                        underline=key[2], overstrike=key[3], size=self.zoomed(spec["size"] * key[4] / DEFAULT_SIZE))
            self.fonts[key] = tkfont.Font(**spec)
        return self.fonts[key]

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
        cv.delete("all")
        if not self.sheet:
            return
        W, H, rh = cv.winfo_width(), cv.winfo_height(), self.rowh
        hw = self.hdr_w()
        ys, y, r = {}, rh, self.top     # row -> (y, height) for visible rows
        while y < H:
            ys[r] = (y, self.rh(r))
            y, r = y + ys[r][1], r + 1
        rows = range(self.top, r)
        cols, x, c = [], hw, self.left
        while x < W:
            w = self.colw(c)
            cols.append((c, x, w))
            x, c = x + w, c + 1
        self.view = (rows, cols)
        if not cols:
            return
        for r in rows:
            y, h = ys[r]
            for c, x, w in cols:
                if (r, c) not in self.merge_at:
                    self.draw_cell(r, c, x, y, w, h, cols)
        first, last = cols[0][0], cols[-1][0]
        for r1, c1, r2, c2 in self.sheet.merges:
            if r2 >= rows.start and r1 < rows.stop and c2 >= first and c1 <= last:
                w = sum(self.colw(i) for i in range(c1, c2 + 1))
                y = self.row_y(r1)
                self.draw_cell(r1, c1, self.col_x(c1), y, w, self.row_y(r2 + 1) - y, None)
        cv.tag_raise("txt")

        # selection outline + active cell box
        xs = {c: (x, w) for c, x, w in cols}

        def box(r1, c1, r2, c2, width, color=None, dash=None):
            if r2 < rows.start or r1 >= rows.stop or c2 < first or c1 > last:
                return
            x1 = xs[c1][0] if c1 >= first else hw - 3
            x2 = sum(xs[c2]) if c2 <= last else W + 3
            y1 = ys[r1][0] if r1 >= self.top else rh - 3
            y2 = sum(ys[r2]) if r2 < rows.stop else H + 3
            cv.create_rectangle(x1, y1, x2, y2, outline=color or C["accent"], width=width, dash=dash,
                                tags="ref" if dash else "")

        sr1, sc1, sr2, sc2 = self.sel()
        if self.multi():
            box(sr1, sc1, sr2, sc2, 1)
        box(*self.span(*self.cur), 2)
        for r1, c1, r2, c2, color in self.ref_marks:   # cells used by the formula being typed
            box(r1, c1, r2, c2, 2, color, (4, 3))
        if self.fill_preview:
            box(*self.fill_preview, 2, "#5f6368", (4, 3))
        # fill handle: small square at the selection's bottom-right corner
        hr1, hc1, hr2, hc2 = (sr1, sc1, sr2, sc2) if self.multi() else self.span(*self.cur)
        self.handle = None
        if hr2 < ALL and hc2 < ALL and hr2 in ys and hc2 in xs and not self.editor:
            hx, hy, s = sum(xs[hc2]), sum(ys[hr2]), max(self.px(3), 3)
            cv.create_rectangle(hx - s, hy - s, hx + s, hy + s, fill=C["accent"], outline=C["cell_bg"], tags="handle")
            self.handle = (hx, hy)

        # headers
        cv.create_rectangle(0, 0, W, rh, fill=C["header_bg"], outline="")
        cv.create_rectangle(0, 0, hw, H, fill=C["header_bg"], outline="")
        hidden_c, hidden_r, mark = self.sheet.hidden_cols, self.sheet.hidden_rows, max(self.px(2), 2)
        for c, x, w in cols:
            if w == 0:
                continue
            on = sc1 <= c <= sc2
            cv.create_rectangle(x, 0, x + w, rh, fill=C["header_sel"] if on else C["header_bg"], outline=C["header_line"])
            cv.create_text(x + w // 2, rh // 2, text=col_name(c), fill=C["header_fg"], font=self.base_font)
            if c - 1 in hidden_c:   # marker: hidden column(s) just left of here
                cv.create_rectangle(x - mark, 2, x + mark, rh - 2, fill=C["hidden_mark"], outline="", tags="hidden")
        for r in rows:
            y, h = ys[r]
            if h == 0:
                continue
            on = sr1 <= r <= sr2
            cv.create_rectangle(0, y, hw, y + h, fill=C["header_sel"] if on else C["header_bg"], outline=C["header_line"])
            cv.create_text(hw - self.px(6), y + h // 2, text=str(r + 1), anchor="e",
                           fill=C["header_fg"], font=self.base_font)
            if r - 1 in hidden_r:   # marker: hidden row(s) just above here
                cv.create_rectangle(2, y - mark, hw - 2, y + mark, fill=C["hidden_mark"], outline="", tags="hidden")
        cv.create_rectangle(0, 0, hw, rh, fill=C["header_bg"], outline=C["header_line"])

        total_r = max(len(self.sheet.rows), self.top + len(rows)) + EXTRA_ROWS
        total_c = max(self.ncols, self.left + len(cols)) + EXTRA_COLS
        self.ys.set(self.top / total_r, (self.top + len(rows)) / total_r)
        self.xs.set(self.left / total_c, (self.left + len(cols)) / total_c)

    def draw_cell(self, r, c, x, y, w, h, cols):
        cv = self.cv
        if w == 0 or h == 0:   # hidden row / column
            return
        st = self.style(r, c)
        bg = st.get("bg") if st else None
        r1, c1, r2, c2 = self.sel()
        if self.multi() and r1 <= r <= r2 and c1 <= c <= c2 and self.span(*self.cur)[:2] != (r, c):
            bg = blend(bg or C["cell_bg"], C["accent"], 0.15)
        cv.create_rectangle(x, y, x + w, y + h, fill=bg or "", outline=C["grid_line"])
        text = self.value(r, c)
        if not text:
            return
        disp = render(text, st.get("numfmt")) if st else text
        is_num = parse_number(text) is not None
        align = (st.get("align") if st else None) or ("right" if is_num else "left")
        font = self.font(st)
        pad = self.px(4)
        avail = w - 2 * pad
        fg = text_color(st.get("fg") if st else None, bg if st and st.get("bg") else None)
        just = "center" if align == "center" else "right" if align == "right" else "left"
        tx = x + w // 2 if just == "center" else x + w - pad if just == "right" else x + pad
        valign = (st.get("valign") if st else None) or "bottom"   # Excel's default
        vpad = self.px(3)
        ty = y + vpad if valign == "top" else y + h // 2 if valign == "center" else y + h - vpad
        # Tk anchor = vertical part (n / - / s) + horizontal part (w / - / e)
        anchor = ({"top": "n", "center": "", "bottom": "s"}[valign]
                  + {"left": "w", "center": "", "right": "e"}[just]) or "center"
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
        cv.create_text(tx, ty, text=disp, anchor=anchor, font=font, fill=fg, tags="txt")

    # ---------- mouse ----------

    def on_press(self, e, extend=False):
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
        if self.on_handle(e.x, e.y):            # start dragging the fill handle
            r1, c1, r2, c2 = self.sel() if self.multi() else self.span(*self.cur)
            self.drag = ("fill", (r1, c1, r2, c2))
            return "break"
        self.commit()
        self.cv.focus_set()
        if not self.sheet:
            return "break"
        b = self.border_at(e.x, e.y)
        if b is not None:
            self.drag = ("resize", b, e.x, self.colw(b))
            return "break"
        r, c = self.hit(e.x, e.y)
        if r < 0 and c < 0:
            return self.select_all()
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
        self.redraw()
        self.app.on_select()
        return "break"

    def on_handle(self, x, y):
        return self.handle is not None and abs(x - self.handle[0]) <= self.px(5) \
            and abs(y - self.handle[1]) <= self.px(5)

    def fill_target(self, x, y):
        """(direction, count) for the fill handle dragged to x, y; None if still inside the block."""
        r1, c1, r2, c2 = self.drag[1]
        W, H = self.cv.winfo_width(), self.cv.winfo_height()
        r, c = self.hit(min(max(x, self.hdr_w()), W - 1), min(max(y, self.rowh), H - 1))
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
        if isinstance(self.drag, tuple) and self.drag[0] == "point":   # drag out a range reference
            W, H = self.cv.winfo_width(), self.cv.winfo_height()
            r, c = self.hit(min(max(e.x, self.hdr_w()), W - 1), min(max(e.y, self.rowh), H - 1))
            self.point_at(self.drag[1], r, c, extend=True)
            return
        if isinstance(self.drag, tuple):  # column resize
            _, c, x0, w0 = self.drag
            self.sheet.px[c] = max(w0 + e.x - x0, self.px(8)) / self.zoom
            self.rowpx = {}    # wrapped text re-flows
            self.redraw()
            return
        W, H, hw, rh = self.cv.winfo_width(), self.cv.winfo_height(), self.hdr_w(), self.rowh
        if self.drag != "cols":   # auto-scroll when dragging past an edge
            if e.y >= H:
                self.top += 1
            elif e.y < rh and self.top > 0:
                self.top -= 1
        if self.drag != "rows":
            if e.x >= W:
                self.left += 1
            elif e.x < hw and self.left > 0:
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
        b = self.border_at(e.x, e.y)
        if b is not None:  # autofit column
            fonts = [self.font(self.style(r, b)) for r in range(min(len(self.sheet.rows), 2000))]
            widths = [f.measure(render(self.value(r, b),(self.style(r, b) or {}).get("numfmt")))
                      for r, f in enumerate(fonts)]
            self.app.set_width(b, self.width_chars((max(widths + [self.px(30)]) + self.px(10)) / self.zoom))
            return "break"
        if e.y >= self.rowh and e.x >= self.hdr_w():
            return self.edit()

    def on_right(self, e):
        """Right-click: keep the selection if clicked inside it, else select there; then the menu."""
        self.commit()
        self.cv.focus_set()
        if not self.sheet:
            return "break"
        r, c = self.hit(e.x, e.y)
        r1, c1, r2, c2 = self.sel()
        inside = (r < 0 or r1 <= r <= r2) and (c < 0 or c1 <= c <= c2)
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
        if self.on_handle(e.x, e.y):
            self.cv["cursor"] = "crosshair"
        else:
            self.cv["cursor"] = "sb_h_double_arrow" if self.border_at(e.x, e.y) is not None else ""

    def on_wheel(self, e, horizontal=False):
        self.commit()
        step = -1 if e.delta > 0 else 1
        if horizontal:
            self.left = max(self.left + step, 0)
        else:
            self.top = max(self.top + 3 * step, 0)
        self.redraw()
        return "break"

    # ---------- keyboard actions ----------

    def on_key(self, e):
        ctrl_only = e.state & 0x4 and not e.state & 0x20000   # Ctrl but not AltGr
        if self.sheet and e.char and e.char.isprintable() and not ctrl_only:
            return self.edit(e.char)

    def clear(self):
        r1, c1, r2, c2 = self.sel_data()
        changes = [(r, c, "") for r in range(r1, r2 + 1) for c in range(c1, c2 + 1) if self.text(r, c)]
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

    def paste(self, mode="all"):
        """mode: 'all', 'values' (shown values, no formats) or 'format' (formats only)."""
        self.commit()
        try:
            data = self.clipboard_get()
        except tk.TclError:
            return "break"
        clip = getattr(self, "clip", None)
        own = clip is not None and clip[0] == data
        if mode == "format" and not own:
            return "break"   # formats only exist for copies made here
        if own and mode != "values":   # our own copy: formulas + formats
            _, sr, sc, rows, cut, styles = clip
        else:
            rows = list(csv.reader(io.StringIO(data.rstrip("\r\n")), delimiter="\t")) or [[""]]
            sr = sc = cut = styles = None
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
                    v = shift(v, r - (sr + i), c - (sc + j))
                if mode != "format":
                    changes.append(("v", r, c, v))
                if styles is not None:
                    changes.append(("s", r, c, styles[i][j]))
        if changes:
            self.app.apply_mixed(changes)
        self.select(r0, c0, r2, c2)
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
        entry.focus_set()
        if hasattr(entry, "fit"):
            entry.fit()
        self.show_refs(entry.get())

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

    def edit(self, initial=None):
        self.commit()
        if not self.sheet:
            return "break"
        r1, c1, r2, c2 = self.span(*self.cur)
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
        e.bind("<FocusOut>", lambda ev: self.after_idle(self.focus_left, e))
        e.bind("<KeyRelease>", lambda ev: (fit(), self.show_refs(e.get())), add="+")
        if initial is not None:  # typing over a cell: arrows finish the edit, like Excel
            for key, d in {"Up": (-1, 0), "Down": (1, 0), "Left": (0, -1), "Right": (0, 1)}.items():
                e.bind(f"<{key}>", lambda ev, d=d: self.commit(*d))
        self.editor = e
        self.show_refs(text)
        return "break"

    def focus_left(self, e):
        """Editor lost focus: finish only if focus went to another control here - switching to
        another program (focus_get() is None) keeps the edit open, like Excel."""
        if self.editor is e and self.focus_get() not in (None, e):
            self.commit()

    def commit(self, dr=0, dc=0, cancel=False):
        e, self.editor = self.editor, None
        if e is None:
            return "break"
        sheet, r, c = e.cell
        val = e.get()
        e.destroy()
        self.ref_marks, self.point = [], None
        if not cancel and sheet is self.sheet and val != self.text(r, c):
            st = self.style(r, c) or {}
            if "\n" in val and not st.get("wrap"):   # a typed line break turns on wrap, like Excel
                self.app.apply_mixed([("v", r, c, val), ("s", r, c, {**st, "wrap": True})])
            else:
                self.app.set_cells([(r, c, val)])
        self.cv.focus_set()
        if dr or dc:
            self.move(dr, dc)
        else:
            self.redraw()
        return "break"
