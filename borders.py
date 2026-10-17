"""Cell borders: which sides each placement sets, the toolbar icons, and the hover popup
(placements + line style + line colour), like Google Sheets."""
import tkinter as tk
from tkinter import colorchooser, ttk

from simplecore.theme import C

from formatting import BORDER_STYLES

PLACEMENTS = [("all", "All borders"), ("inner", "Inner borders"), ("horizontal", "Inner horizontal"),
              ("vertical", "Inner vertical"), ("outer", "Outer borders"), ("left", "Left border"),
              ("top", "Top border"), ("right", "Right border"), ("bottom", "Bottom border"),
              ("none", "Clear borders")]


def sides_for(kind, rect, r, c):
    """Sides of cell (r, c) that placement `kind` draws on the block rect=(r1, c1, r2, c2)."""
    r1, c1, r2, c2 = rect
    top, bottom, left, right = r == r1, r == r2, c == c1, c == c2
    outer = {"top": top, "bottom": bottom, "left": left, "right": right}
    inner_h = {"top": not top, "bottom": not bottom}
    inner_v = {"left": not left, "right": not right}
    want = {
        "all": {"top", "bottom", "left", "right"},
        "outer": {s for s, on in outer.items() if on},
        "inner": {s for s, on in {**inner_h, **inner_v}.items() if on},
        "horizontal": {s for s, on in inner_h.items() if on},
        "vertical": {s for s, on in inner_v.items() if on},
        "top": {"top"} if top else set(),
        "bottom": {"bottom"} if bottom else set(),
        "left": {"left"} if left else set(),
        "right": {"right"} if right else set(),
    }
    return want.get(kind, set())


def restyle(kind, rect, line):
    """-> fn(style, r, c) for App.format_rect: add `line` ([style, colour]) on the chosen sides,
    or remove all borders of the block ('none')."""
    def fn(st, r, c):
        if kind == "none":
            return {**st, "border": None}
        border = dict(st.get("border") or {})
        for side in sides_for(kind, rect, r, c):
            border[side] = list(line)
        return {**st, "border": border or None}
    return fn


# ---------- icons ----------

def placement_icon(bar, kind, img=None):
    """A 2x2 grid: faint dotted lines, with the sides the placement draws in solid ink."""
    s, k = bar.px(16), bar.px(1)
    img = bar.blank(img, s)
    faint, ink = C["muted"], C["icon"]
    a, m, b = k, s // 2, s - 2 * k            # outer edges and middle

    def h(y, color):
        for x in range(a, b + k, 2 * k if color == faint else 1):
            img.put(color, to=(x, y, x + k, y + k))

    def v(x, color):
        for y in range(a, b + k, 2 * k if color == faint else 1):
            img.put(color, to=(x, y, x + k, y + k))
    for y in (a, m, b):
        h(y, faint)
    for x in (a, m, b):
        v(x, faint)
    lines = {"all": "TBLRHV", "inner": "HV", "horizontal": "H", "vertical": "V", "outer": "TBLR",
             "left": "L", "top": "T", "right": "R", "bottom": "B", "none": ""}[kind]
    for code in lines:
        {"T": lambda: h(a, ink), "B": lambda: h(b, ink), "H": lambda: h(m, ink),
         "L": lambda: v(a, ink), "R": lambda: v(b, ink), "V": lambda: v(m, ink)}[code]()
    return img


def line_icon(bar, style, img=None):
    """A short horizontal sample of a line style."""
    w, h = bar.px(28), bar.px(12)
    img = bar.blank(img, w, h)
    ink, k, mid = C["icon"], bar.px(1), h // 2
    if style == "double":
        img.put(ink, to=(0, mid - 2 * k, w, mid - k))
        img.put(ink, to=(0, mid + k, w, mid + 2 * k))
        return img
    thick = 3 * k if style == "thick" else 2 * k if style == "medium" else k
    step = {"dashed": (5 * k, 3 * k), "dotted": (k, k)}.get(style)
    x = 0
    while x < w:
        run = step[0] if step else w
        img.put(ink, to=(x, mid - thick // 2, min(x + run, w), mid - thick // 2 + thick))
        x += run + (step[1] if step else 0)
    return img


# ---------- popup ----------

class BorderButton:
    """Toolbar button; hovering or clicking opens the border popup."""

    def __init__(self, bar, app):
        self.bar, self.app, self.popup = bar, app, None
        self.style = tk.StringVar(bar, value="thin")
        self.color = "#000000"
        for kind, _ in PLACEMENTS:
            bar.icons[f"b-{kind}"] = placement_icon(bar, kind)
        for st in BORDER_STYLES:
            bar.icons[f"line-{st}"] = line_icon(bar, st)
        self.button = ttk.Button(bar, image=bar.icons["b-all"], style="Toolbutton", command=self.toggle)
        self.button.pack(side="left")
        self.button.bind("<Enter>", lambda e: self.open(), add="+")
        self.button.bind("<Leave>", lambda e: self.later_close(), add="+")

    def refresh_icons(self):
        for kind, _ in PLACEMENTS:
            placement_icon(self.bar, kind, self.bar.icons[f"b-{kind}"])
        for st in BORDER_STYLES:
            line_icon(self.bar, st, self.bar.icons[f"line-{st}"])

    def open(self):
        if self.popup is not None:
            return
        from format_bar import Tooltip
        top = tk.Toplevel(self.button)
        top.wm_overrideredirect(True)
        top.attributes("-topmost", True)
        top.configure(background=C["header_line"])
        f = ttk.Frame(top, padding=4)
        f.pack(padx=1, pady=1)
        grid = ttk.Frame(f)
        grid.pack(side="left")
        for i, (kind, tip) in enumerate(PLACEMENTS):
            b = ttk.Button(grid, image=self.bar.icons[f"b-{kind}"], style="Toolbutton",
                           command=lambda k=kind: self.pick(k))
            b.grid(row=i // 5, column=i % 5)
            Tooltip(b, tip)
        ttk.Separator(f, orient="vertical").pack(side="left", fill="y", padx=6)
        opts = ttk.Frame(f)
        opts.pack(side="left")
        self.swatch = tk.Canvas(opts, width=self.bar.px(28), height=self.bar.px(10), highlightthickness=1,
                                highlightbackground=C["header_line"], background=self.color, cursor="hand2")
        self.swatch.grid(row=0, column=0, columnspan=3, pady=(0, 4), sticky="w")
        self.swatch.bind("<Button-1>", lambda e: self.pick_color())
        Tooltip(self.swatch, "Border colour")
        for i, st in enumerate(BORDER_STYLES):
            b = ttk.Radiobutton(opts, image=self.bar.icons[f"line-{st}"], variable=self.style, value=st,
                                style="Toolbutton")
            b.grid(row=1 + i // 3, column=i % 3)
            Tooltip(b, f"{st.capitalize()} line")
        top.update_idletasks()
        x = self.button.winfo_rootx()
        y = self.button.winfo_rooty() + self.button.winfo_height()
        top.wm_geometry(f"+{x}+{y}")
        top.bind("<Leave>", lambda e: self.later_close())
        self.popup = top

    def later_close(self):
        self.button._root().after(300, self.close_if_away)

    def close_if_away(self):
        if self.popup is None or self.choosing:
            return
        px, py = self.button.winfo_pointerxy()
        for w in (self.button, self.popup):
            x, y = w.winfo_rootx(), w.winfo_rooty()
            if x <= px < x + w.winfo_width() and y <= py < y + w.winfo_height():
                return
        self.close()

    choosing = False   # colour dialog open: keep the popup

    def close(self):
        if self.popup is not None:
            self.popup.destroy()
            self.popup = None

    def toggle(self):
        self.close() if self.popup is not None else self.open()

    def pick_color(self):
        self.choosing = True
        try:
            col = colorchooser.askcolor(self.color, title="Border colour", parent=self.button)[1]
        finally:
            self.choosing = False
        if col:
            self.color = col
            if self.popup is not None:
                self.swatch.configure(background=col)

    def pick(self, kind):
        self.close()
        self.apply(kind)

    def apply(self, kind):
        """Apply a placement to the selection with the chosen line style and colour."""
        g = self.app.table
        g.commit()
        if not g.sheet:
            return
        rect = g.sel_cells()
        self.app.format_rect(rect, restyle(kind, rect, [self.style.get(), self.color]))
        g.cv.focus_set()
