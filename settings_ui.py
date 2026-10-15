"""Tools → Settings: language, recovery interval, rows to load, update address."""
import tkinter as tk
from tkinter import messagebox, ttk

import i18n
import recovery
import settings
import sheet_io as io

DEFAULTS = {"lang": "en", "recovery_minutes": 1, "max_rows": 50_000, "update_url": "", "check_updates": False,
            "open_in_tabs": False, "fill_lists": [], "ui_scale": 100, "chart_palette": "office"}
SIZES = [100, 125, 150, 175, 200]
PALETTES = {"office": "Office colours", "colour-blind": "Colour-blind safe (Okabe-Ito)"}
NAMED_FONTS = ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont", "TkFixedFont", "TkTooltipFont",
               "TkCaptionFont", "TkSmallCaptionFont", "TkIconFont")


def get(key):
    return settings.get(key, DEFAULTS[key])


def apply_startup():
    """Settings that take effect when the app starts (also called after saving the dialog)."""
    try:
        recovery.INTERVAL_MS = max(int(float(get("recovery_minutes")) * 60_000), 10_000)
    except (TypeError, ValueError):
        pass
    try:
        io.MAX_ROWS = max(int(get("max_rows")), 1000)
    except (TypeError, ValueError):
        pass
    import charts
    charts.use_palette(get("chart_palette"))


def ui_factor(root):
    """Settings → Interface size: bigger menus, dialogs and text (on the next start). Scales Tk's named
    fonts; returns the factor for the app's own sizes (1.0 = normal)."""
    try:
        factor = int(get("ui_scale")) / 100
    except (TypeError, ValueError):
        factor = 1.0
    if not 0.75 <= factor <= 3:
        factor = 1.0
    if factor != 1.0:
        from tkinter import font as tkfont
        for name in NAMED_FONTS:
            try:
                f = tkfont.nametofont(name, root=root)
                size = f.cget("size")
                f.configure(size=round(size * factor) if size else -round(12 * factor))
            except (tk.TclError, RuntimeError):
                pass
    return factor


class SettingsDialog(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title("Settings")
        self.transient(app)
        self.resizable(False, False)
        f = ttk.Frame(self, padding=12)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text="Language (after a restart)").grid(row=0, column=0, sticky="w", pady=2)
        self.lang = ttk.Combobox(f, values=list(i18n.LANGS.values()), state="readonly", width=22)
        self.lang.set(i18n.LANGS.get(get("lang"), "English"))
        self.lang.grid(row=0, column=1, sticky="w")
        ttk.Label(f, text="Keep unsaved work for recovery every (minutes)").grid(row=1, column=0, sticky="w", pady=2)
        self.minutes = tk.StringVar(master=self, value=str(get("recovery_minutes")))
        ttk.Entry(f, textvariable=self.minutes, width=8).grid(row=1, column=1, sticky="w")
        ttk.Label(f, text="Rows to load per sheet").grid(row=2, column=0, sticky="w", pady=2)
        self.rows = tk.StringVar(master=self, value=str(get("max_rows")))
        ttk.Entry(f, textvariable=self.rows, width=10).grid(row=2, column=1, sticky="w")
        ttk.Label(f, text="Update feed (JSON address)").grid(row=3, column=0, sticky="w", pady=2)
        self.url = tk.StringVar(master=self, value=get("update_url"))
        ttk.Entry(f, textvariable=self.url, width=40).grid(row=3, column=1, sticky="w")
        self.check = tk.BooleanVar(master=self, value=bool(get("check_updates")))
        ttk.Checkbutton(f, text="Check for updates when starting", variable=self.check).grid(
            row=4, column=1, sticky="w")
        self.tabs = tk.BooleanVar(master=self, value=bool(get("open_in_tabs")))
        ttk.Checkbutton(f, text="Open files in tabs in this window (not new windows)", variable=self.tabs).grid(
            row=5, column=1, sticky="w")
        ttk.Label(f, text="Interface size (after a restart)").grid(row=7, column=0, sticky="w", pady=2)
        self.size = ttk.Combobox(f, values=[f"{s}%" for s in SIZES], state="readonly", width=8)
        self.size.set(f"{get('ui_scale')}%")
        self.size.grid(row=7, column=1, sticky="w")
        ttk.Label(f, text="Chart colours").grid(row=8, column=0, sticky="w", pady=2)
        self.palette = ttk.Combobox(f, values=list(PALETTES.values()), state="readonly", width=30)
        self.palette.set(PALETTES.get(get("chart_palette"), PALETTES["office"]))
        self.palette.grid(row=8, column=1, sticky="w")
        ttk.Label(f, text="Fill lists (one per line,\nitems with commas)").grid(row=6, column=0, sticky="nw", pady=2)
        self.lists = tk.Text(f, width=40, height=4)
        self.lists.insert("1.0", "\n".join(", ".join(x) for x in get("fill_lists")))
        self.lists.grid(row=6, column=1, sticky="w", pady=2)
        b = ttk.Frame(f)
        b.grid(row=9, column=0, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(b, text="OK", command=self.ok).pack(side="left", padx=4)
        ttk.Button(b, text="Cancel", command=self.destroy).pack(side="left")
        self.bind("<Escape>", lambda e: self.destroy())

    def fill_lists(self):
        out = []
        for line in self.lists.get("1.0", "end").splitlines():
            items = [x.strip() for x in line.split(",") if x.strip()]
            if len(items) > 1:
                out.append(items)
        return out

    def ok(self):
        try:
            minutes = float(self.minutes.get())
            rows = int(self.rows.get())
            if minutes <= 0 or rows < 1000:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Settings", "Minutes must be above 0 and rows at least 1000.", parent=self)
            return False
        code = next((k for k, v in i18n.LANGS.items() if v == self.lang.get()), "en")
        for k, v in (("lang", code), ("recovery_minutes", minutes), ("max_rows", rows),
                     ("update_url", self.url.get().strip()), ("check_updates", self.check.get()),
                     ("open_in_tabs", self.tabs.get()), ("fill_lists", self.fill_lists()),
                     ("ui_scale", int(self.size.get().rstrip("%") or 100)),
                     ("chart_palette", next((k for k, v in PALETTES.items() if v == self.palette.get()), "office"))):
            settings.set(k, v)
        apply_startup()
        self.app.table.redraw()                 # charts in the new colours
        self.destroy()
        return True
