"""Formatting toolbar: text style preset, size, bold/italic/underline, colors, alignment, clear."""
import tkinter as tk
from tkinter import colorchooser, ttk

from borders import BorderButton
from formatting import DEFAULT_FONT, DEFAULT_SIZE, FONT_KEYS, NUMBER_FORMATS, PRESETS, change_decimals, tint
from theme import C

SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36]
TABLE_COLORS = {"Blue": "4472C4", "Orange": "ED7D31", "Green": "70AD47", "Gold": "FFC000",
                "Sky": "5B9BD5", "Gray": "A5A5A5", "Dark": "404040"}


def table_colors(color, light):
    """-> (header fill, band fill) for a table style."""
    if light:
        return tint(color, 0.6), tint(color, 0.85)
    return "#" + color.lower(), tint(color, 0.8)


class Tooltip:
    def __init__(self, widget, text):
        self.widget, self.text, self.tip = widget, text, None
        widget.bind("<Enter>", self.show, add="+")
        widget.bind("<Leave>", self.hide, add="+")
        widget.bind("<ButtonPress>", self.hide, add="+")

    def show(self, _=None):
        if self.tip:
            return
        x = self.widget.winfo_rootx()
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 2
        self.tip = tk.Toplevel(self.widget)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f"+{x}+{y}")
        import i18n
        import platform_ui
        text = platform_ui.mac_keys(i18n.tr(self.text))   # macOS: ⌘B, not Ctrl+B
        tk.Label(self.tip, text=text, bg=C["tooltip_bg"], fg=C["tooltip_fg"], relief="solid", bd=1,
                 padx=4).pack()

    def hide(self, _=None):
        if self.tip:
            self.tip.destroy()
            self.tip = None


COMMON_FONTS = ["Calibri", "Arial", "Cambria", "Candara", "Comic Sans MS", "Consolas", "Courier New",
                "Georgia", "Segoe UI", "Tahoma", "Times New Roman", "Trebuchet MS", "Verdana"]


def font_choices(widget):
    """Common fonts that are installed here."""
    from tkinter import font as tkfont
    have = {f.lower() for f in tkfont.families(widget)}
    return [f for f in COMMON_FONTS if f.lower() in have] or [DEFAULT_FONT]


VALIGNS = ("top", "center", "bottom")
VALIGN_TIPS = {"top": "Align top", "center": "Align middle", "bottom": "Align bottom"}


class IconDropdown:
    """A toolbar button showing the current choice; hovering (or clicking) it opens the choices
    in a small strip underneath. Picking one calls on_pick(value)."""

    def __init__(self, bar, variable, options, icons, on_pick, default):
        self.var, self.options, self.icons = variable, options, icons
        self.on_pick, self.default, self.popup = on_pick, default, None
        self.button = ttk.Button(bar, image=icons[default], style="Toolbutton", command=self.toggle)
        self.button.pack(side="left")
        self.button.bind("<Enter>", lambda e: self.open(), add="+")
        self.button.bind("<Leave>", lambda e: self.later_close(), add="+")
        variable.trace_add("write", lambda *a: self.show_current())

    def show_current(self):
        self.button.configure(image=self.icons.get(self.var.get()) or self.icons[self.default])

    def open(self):
        if self.popup is not None:
            return
        top = tk.Toplevel(self.button)
        top.wm_overrideredirect(True)
        top.attributes("-topmost", True)     # never hidden behind the main window
        top.configure(background=C["header_line"])
        frame = ttk.Frame(top, padding=2)
        frame.pack(padx=1, pady=1)
        for value, tip in self.options:
            b = ttk.Radiobutton(frame, image=self.icons[value], variable=self.var, value=value,
                                style="Toolbutton", command=lambda v=value: self.pick(v))
            b.pack(side="left")
            Tooltip(b, tip)
        x = self.button.winfo_rootx()
        y = self.button.winfo_rooty() + self.button.winfo_height()
        top.wm_geometry(f"+{x}+{y}")
        top.bind("<Leave>", lambda e: self.later_close())
        self.popup = top

    def later_close(self):
        self.button._root().after(300, self.close_if_away)

    def close_if_away(self):
        """Close unless the mouse is over the button or the strip."""
        if self.popup is None:
            return
        px, py = self.button.winfo_pointerxy()
        for w in (self.button, self.popup):
            x, y = w.winfo_rootx(), w.winfo_rooty()
            if x <= px < x + w.winfo_width() and y <= py < y + w.winfo_height():
                return
        self.close()

    def close(self):
        if self.popup is not None:
            self.popup.destroy()
            self.popup = None

    def toggle(self):
        self.close() if self.popup is not None else self.open()

    def pick(self, value):
        self.close()
        self.on_pick(value)


class FormatBar(ttk.Frame):
    """app must provide: format_cells(fn) where fn(style dict) -> new style dict, and table (Grid)."""

    def __init__(self, master, app, scale):
        super().__init__(master, padding=(4, 0, 4, 4))
        self.app, self.scale = app, scale
        self.syncing = False
        self.icons = {}


        self.preset = ttk.Combobox(self, values=list(PRESETS), width=11, state="readonly")
        self.preset.pack(side="left")
        self.preset.bind("<<ComboboxSelected>>", lambda e: self.set_preset())
        Tooltip(self.preset, "Text style")
        self.font = ttk.Combobox(self, values=font_choices(self), width=14)
        self.font.pack(side="left", padx=(4, 0))
        self.font.bind("<<ComboboxSelected>>", lambda e: self.set_font())
        self.font.bind("<Return>", lambda e: self.set_font())
        Tooltip(self.font, "Font")
        self.size = ttk.Combobox(self, values=SIZES, width=4)
        self.size.pack(side="left", padx=(4, 0))
        self.size.bind("<<ComboboxSelected>>", lambda e: self.set_size())
        self.size.bind("<Return>", lambda e: self.set_size())
        Tooltip(self.size, "Font size")
        for text, step, tip in (("A▴", 1, "Bigger text (Ctrl+Shift+>)"), ("A▾", -1, "Smaller text (Ctrl+Shift+<)")):
            b = ttk.Button(self, text=text, style="Toolbutton", width=3, command=lambda s=step: self.step_size(s))
            b.pack(side="left")
            Tooltip(b, tip)
        self.sep()

        self.vars = {}
        for key, label, tip in (("bold", "B", "Bold (Ctrl+B)"), ("italic", "I", "Italic (Ctrl+I)"),
                                ("underline", "U", "Underline (Ctrl+U)"), ("strike", "S", "Strikethrough (Ctrl+5)")):
            v = self.vars[key] = tk.BooleanVar(self)   # explicit master: never another Tk instance
            b = ttk.Checkbutton(self, text=label, variable=v, style=f"{label}.Toolbutton", width=2,
                                command=lambda k=key: self.toggle(k))
            b.pack(side="left")
            Tooltip(b, tip)
        b = ttk.Button(self, text="N", style="N.Toolbutton", width=2, command=self.normal)
        b.pack(side="left")
        Tooltip(b, "Normal (remove bold, italic, underline)")
        self.sep()

        self.fg_btn = self.icon_button("fg", "Text color", lambda: self.pick("fg"))
        self.bg_btn = self.icon_button("bg", "Fill color", lambda: self.pick("bg"))
        self.sep()

        self.align = tk.StringVar(self)
        for a in ("left", "center", "right", "justify"):
            self.icons[a] = self.align_icon(a)
        self.align_menu = IconDropdown(self, self.align, [(a, f"Align {a}") for a in ("left", "center", "right", "justify")],
                                       {a: self.icons[a] for a in ("left", "center", "right", "justify")},
                                       self.set_align, default="left")
        self.valign = tk.StringVar(self)
        for v in VALIGNS:
            self.icons[f"v-{v}"] = self.valign_icon(v)
        self.valign_menu = IconDropdown(self, self.valign, [(v, VALIGN_TIPS[v]) for v in VALIGNS],
                                        {v: self.icons[f"v-{v}"] for v in VALIGNS},
                                        self.set_valign, default="bottom")
        self.sep()

        for text, delta, tip in (("⇤", -1, "Decrease indent"), ("⇥", 1, "Increase indent")):
            b = ttk.Button(self, text=text, style="Toolbutton", width=2, command=lambda d=delta: self.indent(d))
            b.pack(side="left")
            Tooltip(b, tip)
        rot = ttk.Menubutton(self, text="⟲", style="Toolbutton", width=2)
        rot.pack(side="left")
        Tooltip(rot, "Rotate text")
        m = tk.Menu(rot, tearoff=False)
        for label, deg in (("None", 0), ("Angle up (45°)", 45), ("Angle down (−45°)", -45), ("Up (90°)", 90),
                           ("Down (−90°)", -90)):
            m.add_command(label=label, command=lambda d=deg: self.rotate(d))
        rot["menu"] = m
        self.rotate_menu = m
        self.wrap = tk.BooleanVar(self)
        img = self.icons["wrap"] = self.wrap_icon()
        b = ttk.Checkbutton(self, image=img, variable=self.wrap, style="Toolbutton", command=self.toggle_wrap)
        b.pack(side="left", padx=(4, 0))
        Tooltip(b, "Wrap text (show long text on several lines)")

        self.merged = tk.BooleanVar(self)
        img = self.icons["merge"] = self.merge_icon()
        b = ttk.Checkbutton(self, image=img, variable=self.merged, style="Toolbutton", command=self.toggle_merge)
        b.pack(side="left", padx=(4, 0))
        Tooltip(b, "Merge cells / unmerge")
        self.borders = BorderButton(self, app)
        self.borders.button.pack_configure(padx=(4, 0))
        self.sep()

        nf = ttk.Menubutton(self, text="123", style="Toolbutton")
        nf.pack(side="left")
        Tooltip(nf, "Number format")
        self.numfmt_menu = tk.Menu(nf, tearoff=False)
        for label, code in NUMBER_FORMATS:
            self.numfmt_menu.add_command(label=label, command=lambda c=code: self.set_numfmt(c))
        self.numfmt_menu.add_separator()
        self.numfmt_menu.add_command(label="Custom…", command=self.custom_numfmt)
        nf["menu"] = self.numfmt_menu
        for text, delta, tip in ((".0", -1, "Fewer decimal places"), (".00", 1, "More decimal places")):
            b = ttk.Button(self, text=text, style="Toolbutton", width=4, command=lambda d=delta: self.decimals(d))
            b.pack(side="left")
            Tooltip(b, tip)
        self.sep()

        b = ttk.Button(self, text="✕ Clear", style="Toolbutton", command=self.clear)
        b.pack(side="left")
        Tooltip(b, "Clear formatting")
        # a drawn brush: the 🖌 emoji made Windows load the colour-emoji font (~0.3 s at every start)
        self.icons["painter"] = self.painter_icon()
        self.painter_btn = ttk.Button(self, image=self.icons["painter"], style="Toolbutton",
                                      command=self.app.format_painter)
        self.painter_btn.pack(side="left")
        Tooltip(self.painter_btn, "Format painter: copy this formatting, then select where it goes (Esc cancels)")
        self.sep()
        self.table_menu()
        styles = ttk.Menubutton(self, text="Styles", style="Toolbutton")
        styles.pack(side="left")
        Tooltip(styles, "Cell styles (Good, Bad, Heading 1…)")
        m = tk.Menu(styles, tearoff=False)
        from formatting import CELL_STYLES
        for name in CELL_STYLES:
            m.add_command(label=name, command=lambda n=name: self.cell_style(n))
        styles["menu"] = m
        self.styles_menu = m
        self.colors = {"fg": "#000000", "bg": "#ffff00"}   # last used, applied by the icon click
        self.draw_color_icons()
        self.setup_flow()

    # ---------- layout ----------

    def setup_flow(self):
        """Lay the buttons out in rows: on a narrow window (MacBook Air, half a screen) the groups between
        separators that don't fit go behind a » button, which shows them as more rows (and hides them again)."""
        import settings
        self.groups, group = [], []
        for w in self.pack_slaves():
            info = w.pack_info()
            pad = info.get("padx", 0)
            pad = tuple(int(p) for p in pad) if isinstance(pad, (tuple, list)) else (int(pad), int(pad))
            if isinstance(w, ttk.Separator):
                self.groups.append(group)
                group = []
            else:
                group.append((w, pad if len(pad) == 2 else (pad[0], pad[0])))
            w.pack_forget()
        self.groups.append(group)
        self.seps = [w for w in self.winfo_children() if isinstance(w, ttk.Separator)]
        self.expanded = bool(settings.get("toolbar_expanded", False))
        self.more_btn = ttk.Button(self, text="»", style="Toolbutton", width=2, command=self.toggle_more)
        self.more_tip = Tooltip(self.more_btn, "More tools")
        self.flow_rows, self.hidden_groups = 1, 0
        self._flow_key = None
        # one row high from the start, nothing placed until the real width is known: no jump at the first open
        self.configure(height=self.row_height() + 4, width=1)
        self.bind("<Configure>", lambda e: self.flow(), add="+")
        for g in self.groups:                              # a button got wider (language, icons): lay out again
            for w, _ in g:
                w.bind("<Configure>", lambda e: self.after_idle(self.flow), add="+")

    def row_height(self):
        return max((w.winfo_reqheight() for g in self.groups for w, _ in g), default=1)

    def toggle_more(self):
        import settings
        self.expanded = not self.expanded
        settings.set("toolbar_expanded", self.expanded)
        self.flow()

    def flow(self, width=None):
        """Place the groups left to right. When they don't all fit on one row, a » button at its right end
        shows the rest as more rows (expanded) or keeps them hidden (one row: more room for the sheet)."""
        width = width or self.winfo_width()
        if width <= 1:                                     # not mapped yet: wait for the real width
            return
        sep_w = 14                                         # padx 6 + line + 6, as the packed separators had
        row_h = self.row_height()
        widths = [sum(w.winfo_reqwidth() + p[0] + p[1] for w, p in g) for g in self.groups]
        key = (width, row_h, tuple(widths), self.expanded)
        if key == self._flow_key:
            return
        self._flow_key = key
        pad_l, pad_b = 4, 4                                # the frame's padding (4, 0, 4, 4)
        gap = self.px(2)
        more_w = self.more_btn.winfo_reqwidth() + 6
        fits = sum(widths) + sep_w * (len(widths) - 1) <= width - 2 * pad_l
        x, row = pad_l, 0
        spots = []                                         # (group index, x, row)
        for i, gw in enumerate(widths):
            room = width - pad_l - (0 if fits or row else more_w)   # the first row keeps room for »
            if spots and x + sep_w + gw > room:            # doesn't fit: next row
                x, row = pad_l, row + 1
            elif spots:
                x += sep_w
            spots.append((i, x, row))
            x += gw
        shown = spots if self.expanded or fits else [s for s in spots if s[2] == 0]
        self.hidden_groups = len(spots) - len(shown)
        for sep in self.seps:
            sep.place_forget()
        for n, (i, x, row) in enumerate(spots):
            if n >= len(shown):
                for w, _ in self.groups[i]:
                    w.place_forget()
                continue
            y = row * (row_h + gap)
            if n and spots[n - 1][2] == row and n - 1 < len(self.seps):   # separator before it on the same row
                self.seps[n - 1].place(x=x - sep_w + 6, y=y, width=2, height=row_h)
            for w, (pl, pr) in self.groups[i]:
                x += pl
                w.place(x=x, y=y + (row_h - w.winfo_reqheight()) // 2)
                x += w.winfo_reqwidth() + pr
        if fits:
            self.more_btn.place_forget()
        else:
            self.more_btn.configure(text="«" if self.expanded else "»")
            self.more_tip.text = "Fewer tools" if self.expanded else "More tools"
            self.more_btn.place(x=width - pad_l - self.more_btn.winfo_reqwidth(),
                                y=(row_h - self.more_btn.winfo_reqheight()) // 2)
        self.flow_rows = (shown[-1][2] + 1) if shown else 1
        total = self.flow_rows * row_h + (self.flow_rows - 1) * gap + pad_b
        self.configure(height=total, width=1)              # takes the width it's given, asks for none

    # ---------- widgets ----------

    def sep(self):
        ttk.Separator(self, orient="vertical").pack(side="left", fill="y", padx=6)

    def px(self, n):
        return max(int(n * self.scale), 1)

    def icon_button(self, kind, tip, cmd):
        f = ttk.Frame(self)
        f.pack(side="left")
        b = ttk.Button(f, style="Toolbutton", command=cmd)
        b.pack(side="left")
        Tooltip(b, tip + " (click to choose)")
        return b

    def draw_color_icons(self):
        """'A' over a color bar for text, a filled square for fill."""
        s = self.px(16)
        for kind in ("fg", "bg"):
            img = tk.PhotoImage(master=self, width=s, height=s)
            col = self.colors[kind]
            bar = self.px(4)
            if kind == "fg":
                # letter A from lines
                for i in range(s - bar - self.px(2)):
                    t = i / (s - bar - self.px(2))
                    xl = int(s / 2 - t * s * 0.35)
                    xr = int(s / 2 + t * s * 0.35)
                    img.put(C["icon"], to=(xl, i, xl + self.px(2), i + 1))
                    img.put(C["icon"], to=(xr - self.px(2), i, xr, i + 1))
                mid = int((s - bar) * 0.6)
                img.put(C["icon"], to=(int(s * 0.32), mid, int(s * 0.68), mid + self.px(2)))
            else:
                img.put(C["muted"], to=(self.px(2), self.px(1), s - self.px(2), s - bar - self.px(1)))
                img.put(C["cell_bg"], to=(self.px(3), self.px(2), s - self.px(3), s - bar - self.px(2)))
            img.put(col, to=(0, s - bar, s, s))
            self.icons[kind] = img
        self.fg_btn["image"] = self.icons["fg"]
        self.bg_btn["image"] = self.icons["bg"]

    def blank(self, img, w, h=None):
        """A new image, or the given one cleared (so buttons using it update in place)."""
        if img is None:
            return tk.PhotoImage(master=self, width=w, height=h or w)
        img.blank()
        return img

    def align_icon(self, kind, img=None):
        s = self.px(16)
        img = self.blank(img, s)
        lens = [1.0, 0.6, 1.0, 0.6] if kind != "justify" else [1.0, 1.0, 1.0, 0.6]
        th = self.px(2)
        for i, ln in enumerate(lens):
            w = int((s - self.px(2)) * ln)
            y = self.px(2) + i * int(s / 4.2)
            if kind == "right":
                x = s - self.px(1) - w
            elif kind == "center":
                x = (s - w) // 2
            else:
                x = self.px(1)
            img.put(C["icon"], to=(x, y, x + w, y + th))
        return img

    def valign_icon(self, kind, img=None):
        """A text line (thick bar) at the top / middle / bottom between two thin edges."""
        s, k = self.px(16), self.px(1)
        img = self.blank(img, s)
        ink = C["icon"]
        img.put(C["muted"], to=(0, 0, s, k))
        img.put(C["muted"], to=(0, s - k, s, s))
        bar = 2 * k
        y = {"top": 3 * k, "center": (s - bar) // 2, "bottom": s - 3 * k - bar}[kind]
        img.put(ink, to=(3 * k, y, s - 3 * k, y + bar))
        mid = s // 2
        if kind != "center":   # small arrow pointing at the bar
            up = kind == "top"
            tip = y + bar + k if up else y - k - 1
            for i in range(3 * k):
                row = tip + i if up else tip - i
                img.put(ink, to=(mid - i, row, mid + i + 1, row + 1))
        return img

    def wrap_icon(self, img=None):
        """Google-Sheets style: a line that runs to the edge and turns back underneath."""
        s, k = self.px(16), self.px(1)
        img = self.blank(img, s)
        dark = C["icon"]
        img.put(dark, to=(0, 0, k, s))                         # left edge
        img.put(dark, to=(s - k, 0, s, s))                     # right edge
        y1, y2 = 4 * k, 10 * k
        img.put(dark, to=(3 * k, y1, 12 * k, y1 + k))           # line to the right ...
        img.put(dark, to=(11 * k, y1, 12 * k, y2 + k))          # ... down ...
        img.put(dark, to=(5 * k, y2, 12 * k, y2 + k))           # ... and back left
        for i in range(3 * k):                                 # arrow head pointing left
            img.put(dark, to=(5 * k + i, y2 - i, 5 * k + i + 1, y2 + k + i))
        return img

    def refresh_icons(self):
        """Redraw icons in the current theme's ink (after a light/dark switch)."""
        for a in ("left", "center", "right", "justify"):
            self.align_icon(a, self.icons[a])
        for v in VALIGNS:
            self.valign_icon(v, self.icons[f"v-{v}"])
        self.wrap_icon(self.icons["wrap"])
        self.merge_icon(self.icons["merge"])
        self.painter_icon(self.icons["painter"])
        self.borders.refresh_icons()
        self.draw_color_icons()

    def toggle_wrap(self):
        on = self.wrap.get()
        self.app.format_cells(lambda st: {**st, "wrap": on})
        self.done()

    def merge_icon(self, img=None):
        """A box with arrows pointing outwards: <-|->."""
        s, k = self.px(16), self.px(1)
        img = self.blank(img, s)
        dark = C["icon"]
        for edge in ((0, 0, s, k), (0, s - k, s, s), (0, 0, k, s), (s - k, 0, s, s)):
            img.put(dark, to=edge)
        mid = s // 2
        img.put(dark, to=(3 * k, mid, s - 3 * k, mid + k))   # shaft
        for i in range(3 * k):   # arrow heads
            img.put(dark, to=(3 * k + i, mid - i, 3 * k + i + 1, mid + i + 1))
            img.put(dark, to=(s - 3 * k - i - 1, mid - i, s - 3 * k - i, mid + i + 1))
        return img

    def painter_icon(self, img=None):
        """A paint brush: bristles (accent colour), metal band, handle."""
        s, k = self.px(16), self.px(1)
        img = self.blank(img, s)
        ink, tip = C["icon"], C["accent"]
        img.put(tip, to=(3 * k, k, 13 * k, 6 * k))                 # bristles
        img.put(ink, to=(3 * k, 6 * k, 13 * k, 8 * k))             # band
        img.put(ink, to=(7 * k, 8 * k, 9 * k, 15 * k))             # handle
        return img

    def toggle_merge(self):
        if self.merged.get():
            if not self.app.merge():
                self.merged.set(False)
        else:
            self.app.unmerge()
        self.done()

    def table_icon(self, color, light):
        """Mini table preview: header bar + banded rows."""
        head, band = table_colors(color, light)
        w, h, rows = self.px(30), self.px(18), 4
        rh = h // (rows + 1)
        img = tk.PhotoImage(master=self, width=w, height=h)
        img.put("#ffffff", to=(0, 0, w, h))
        img.put(head, to=(0, 0, w, rh))
        for i in range(rows):
            if i % 2 == 0:
                img.put(band, to=(0, rh * (i + 1), w, rh * (i + 2)))
        for edge in ((0, 0, w, 1), (0, h - 1, w, h), (0, 0, 1, h), (w - 1, 0, w, h)):
            img.put("#9a9a9a", to=edge)
        return img

    def table_menu(self):
        mb = ttk.Menubutton(self, text="Table", style="Toolbutton")
        mb.pack(side="left")
        Tooltip(mb, "Format as table (selection, or the data block around the cursor)")
        menu = tk.Menu(mb, tearoff=False)
        for light in (False, True):
            menu.add_command(label="Light" if light else "Medium", state="disabled")
            for name, color in TABLE_COLORS.items():
                img = self.icons[f"table-{name}-{light}"] = self.table_icon(color, light)
                menu.add_command(label=f"  {name}", image=img, compound="left",
                                 command=lambda c=color, lt=light: self.apply_table(c, lt))
            menu.add_separator()
        menu.add_command(label="Remove table style", command=self.remove_table)
        mb["menu"] = menu
        self.icons["table"] = self.table_icon(TABLE_COLORS["Blue"], False)
        mb.configure(image=self.icons["table"], compound="left")

    def table_rect(self):
        g = self.app.table
        g.commit()
        return g.sel_cells() if g.multi() else g.region(*g.cur)

    def apply_table(self, color, light):
        g = self.app.table
        if not g.sheet:
            return
        r1, c1, r2, c2 = rect = self.table_rect()
        head, band = table_colors(color, light)
        head_fg = "#000000" if light else "#ffffff"

        def fn(st, r, c):
            if r == r1:   # header row
                return {**st, "bg": head, "fg": head_fg, "bold": True}
            st = {k: v for k, v in st.items() if k != "bg"}
            return {**st, "bg": band} if (r - r1 - 1) % 2 == 0 else st

        self.app.format_rect(rect, fn)
        g.select(r1, c1, r2, c2)
        self.done()

    def remove_table(self):
        g = self.app.table
        if not g.sheet:
            return
        r1, c1, r2, c2 = rect = self.table_rect()
        self.app.format_rect(rect, lambda st, r, c: {
            k: v for k, v in st.items() if k != "bg" and not (r == r1 and k in ("fg", "bold"))})
        g.select(r1, c1, r2, c2)
        self.done()

    # ---------- state ----------

    def sync(self, st):
        """Show the active cell's format on the buttons."""
        st = st or {}
        self.syncing = True
        for k, v in self.vars.items():
            v.set(bool(st.get(k)))
        self.align.set(st.get("align", ""))
        self.font.set(st.get("font", DEFAULT_FONT))
        self.valign.set(st.get("valign", "bottom"))
        self.wrap.set(bool(st.get("wrap")))
        self.size.set(f"{st.get('size', DEFAULT_SIZE):g}")
        font = {k: st[k] for k in FONT_KEYS if k in st}
        self.preset.set(next((n for n, p in PRESETS.items() if p == font), ""))
        self.syncing = False

    # ---------- actions ----------

    def done(self):
        self.app.table.cv.focus_set()

    def toggle(self, key):
        on = self.vars[key].get()
        self.app.format_cells(lambda st: {**st, key: on})
        self.done()

    def normal(self):
        self.app.format_cells(lambda st: {k: v for k, v in st.items()
                                          if k not in ("bold", "italic", "underline", "strike")})
        self.done()

    def set_size(self):
        try:
            size = float(self.size.get())
        except ValueError:
            return
        if 1 <= size <= 409:
            self.app.format_cells(lambda st: {**st, "size": size})
        self.done()

    def step_size(self, step):
        """A▴ / A▾: each selected cell to the next bigger / smaller standard size (like Excel)."""
        def bump(st):
            cur = float(st.get("size") or DEFAULT_SIZE)
            if step > 0:
                new = next((s for s in SIZES if s > cur), min(cur + 12, 409))
            else:
                new = next((s for s in reversed(SIZES) if s < cur), max(cur - 1, 1))
            return {**st, "size": float(new)}
        self.app.format_cells(bump)
        self.done()

    def set_preset(self):
        p = PRESETS[self.preset.get()]
        self.app.format_cells(lambda st: {**{k: v for k, v in st.items() if k not in FONT_KEYS}, **p})
        self.done()

    def indent(self, delta):
        """Indent the text in from the edge (Excel indent levels 0-15)."""
        self.app.format_cells(lambda st: {**st, "indent": max(0, min(15, int(st.get("indent", 0)) + delta)) or None})
        self.done()

    def rotate(self, degrees):
        """Text at an angle (degrees, counter-clockwise); 0 = straight."""
        self.app.format_cells(lambda st: {**st, "rotate": degrees or None})
        self.done()

    def cell_style(self, name):
        """An Excel cell style: replaces the look (fonts, fill, borders), keeps number format and alignment."""
        from formatting import CELL_STYLES
        look = CELL_STYLES[name]
        keep = ("numfmt", "align", "valign", "wrap", "unlocked")
        self.app.format_cells(lambda st: {**{k: v for k, v in st.items() if k in keep}, **look})
        self.done()

    def set_numfmt(self, code):
        self.app.format_cells(lambda st: {**st, "numfmt": code})
        self.done()

    def custom_numfmt(self):
        from tkinter import simpledialog
        g = self.app.table
        cur = (g.style(*g.cur) or {}).get("numfmt", "General")
        code = simpledialog.askstring("Number format", "Format code (like Excel), e.g. 0.000  or  #,##0 \"kg\":",
                                      initialvalue=cur, parent=self)
        if code:
            self.set_numfmt(None if code.strip() == "General" else code.strip())

    def decimals(self, delta):
        """Each selected cell gets one decimal more / fewer (based on its own format and value)."""
        g = self.app.table

        def fn(st, r, c):
            new = change_decimals(st.get("numfmt"), delta, g.value(r, c))
            return {**st, "numfmt": new if new != "General" else None}
        g.commit()
        self.app.format_rect(g.cells_rects(), fn)
        self.done()

    def set_font(self):
        name = self.font.get().strip()
        if name:
            self.app.format_cells(lambda st: {**st, "font": name})
        self.done()

    def set_align(self, a):
        self.app.format_cells(lambda st: {**st, "align": a})
        self.done()

    def set_valign(self, v):
        """Bottom is the default (like Excel), so it isn't stored."""
        self.app.format_cells(lambda st: {**st, "valign": None if v == "bottom" else v})
        self.done()

    def pick(self, kind):
        title = "Text color" if kind == "fg" else "Fill color"
        col = colorchooser.askcolor(self.colors[kind], title=title, parent=self)[1]
        if col:
            self.colors[kind] = col
            self.draw_color_icons()
            self.app.format_cells(lambda st: {**st, kind: col})
        self.done()

    def clear(self):
        self.app.format_cells(lambda st: {k: v for k, v in st.items() if k == "numfmt"})
        self.done()

    def shortcut(self, key):
        """Ctrl+B / I / U from the grid."""
        self.vars[key].set(not self.vars[key].get())
        self.toggle(key)
        return "break"
