"""Light / dark colours for the whole app. `C` always holds the active palette."""
import sys
import tkinter as tk
from tkinter import ttk

LIGHT = {
    "dark": False,
    "bg": "#f0f0f0", "fg": "#000000", "muted": "#555555",
    "cell_bg": "#ffffff", "cell_fg": "#000000", "grid_line": "#e3e3e3",
    "header_bg": "#f3f3f3", "header_sel": "#d3e3fd", "header_fg": "#444444", "header_line": "#d0d0d0",
    "accent": "#1a73e8", "hidden_mark": "#8a8a8a", "icon": "#333333",
    "field_bg": "#ffffff", "select_bg": "#cce0ff", "menu_bg": "#f9f9f9", "menu_active": "#d3e3fd",
    "tooltip_bg": "#ffffe0", "tooltip_fg": "#000000",
}
DARK = {
    "dark": True,
    "bg": "#202124", "fg": "#e8eaed", "muted": "#9aa0a6",
    "cell_bg": "#1e1f22", "cell_fg": "#e8eaed", "grid_line": "#3a3b3e",
    "header_bg": "#2b2c2f", "header_sel": "#3c4a63", "header_fg": "#bdc1c6", "header_line": "#45464a",
    "accent": "#8ab4f8", "hidden_mark": "#9aa0a6", "icon": "#e8eaed",
    "field_bg": "#2b2c2f", "select_bg": "#3c4a63", "menu_bg": "#2b2c2f", "menu_active": "#3c4a63",
    "tooltip_bg": "#3c4043", "tooltip_fg": "#e8eaed",
}
HIGH = {   # high contrast: white on black, yellow accents, strong lines
    "dark": True, "contrast": True,
    "bg": "#000000", "fg": "#ffffff", "muted": "#d0d0d0",
    "cell_bg": "#000000", "cell_fg": "#ffffff", "grid_line": "#9a9a9a",
    "header_bg": "#000000", "header_sel": "#00309c", "header_fg": "#ffffff", "header_line": "#ffffff",
    "accent": "#ffff00", "hidden_mark": "#ffff00", "icon": "#ffffff",
    "field_bg": "#000000", "select_bg": "#00309c", "menu_bg": "#000000", "menu_active": "#00309c",
    "tooltip_bg": "#000000", "tooltip_fg": "#ffff00",
}
C = dict(LIGHT)
LIGHT_THEME = None   # the platform's native ttk theme, remembered on first use


def text_color(fg, bg):
    """Colour for cell text: the cell's own colour, unless it would vanish on a dark empty cell."""
    if bg:                                   # filled cells keep light-mode contrast
        return fg or "#000000"
    if not fg:
        return C["cell_fg"]
    return C["cell_fg"] if C["dark"] and near_black(fg) else fg


def near_black(hex_color):
    """Black / very dark grey: invisible on a dark cell. Colours like red or navy stay."""
    return max(int(hex_color[i:i + 2], 16) for i in (1, 3, 5)) < 0x50


def base_styles(style, scale):
    """App-specific ttk styles; re-applied after every theme switch (they are per-theme)."""
    style.configure("Bottom.TNotebook", tabposition="sw")
    style.configure("Bottom.TNotebook.Tab", padding=(int(14 * scale), int(3 * scale)))
    style.configure("Icon.Toolbutton", font=("Segoe UI Symbol", 12))
    for name, font in (("B", ("Segoe UI", 10, "bold")), ("I", ("Segoe UI", 10, "italic")),
                       ("U", ("Segoe UI", 10, "underline")), ("S", ("Segoe UI", 10, "overstrike")),
                       ("N", ("Segoe UI", 10))):
        style.configure(f"{name}.Toolbutton", font=font)


def apply(app, dark, contrast=False):
    """Switch the whole window between light, dark and high contrast."""
    global LIGHT_THEME
    C.clear()
    C.update(HIGH if contrast else DARK if dark else LIGHT)
    dark = C["dark"]
    style = ttk.Style(app)
    if LIGHT_THEME is None:
        LIGHT_THEME = style.theme_use()
    if dark:
        style.theme_use("clam")              # native Windows widgets can't be recoloured
        bg, fg, field = C["bg"], C["fg"], C["field_bg"]
        style.configure(".", background=bg, foreground=fg, fieldbackground=field, bordercolor=C["header_line"],
                        darkcolor=bg, lightcolor=bg, troughcolor=C["cell_bg"], insertcolor=fg,
                        selectbackground=C["select_bg"], selectforeground=fg, arrowcolor=fg)
        style.map(".", background=[("disabled", bg), ("active", C["menu_active"])],
                  foreground=[("disabled", C["muted"])])
        style.map("Toolbutton", background=[("selected", C["select_bg"]), ("active", C["menu_active"])])
        style.map("TCombobox", fieldbackground=[("readonly", field)], foreground=[("readonly", fg)],
                  selectbackground=[("readonly", field)], selectforeground=[("readonly", fg)])
        style.configure("TNotebook", background=bg, bordercolor=C["header_line"])
        style.configure("TNotebook.Tab", background=C["header_bg"], foreground=C["muted"])
        style.map("TNotebook.Tab", background=[("selected", C["cell_bg"])], foreground=[("selected", fg)])
        style.configure("TScrollbar", background=C["header_bg"], troughcolor=C["cell_bg"], arrowcolor=fg)
        app.option_add("*TCombobox*Listbox.background", field)
        app.option_add("*TCombobox*Listbox.foreground", fg)
    else:
        style.theme_use(LIGHT_THEME)
    base_styles(style, app.scale)
    app.configure(background=C["bg"])
    _recolor(app)
    _title_bar(app, dark)


def _recolor(widget):
    """Plain tk widgets (menus, canvas, text, tooltips) don't follow ttk themes: set them directly."""
    for w in [widget] + _descendants(widget):
        cls = w.winfo_class()
        try:
            if cls == "Menu":
                w.configure(background=C["menu_bg"], foreground=C["fg"], activebackground=C["menu_active"],
                            activeforeground=C["fg"], disabledforeground=C["muted"], selectcolor=C["fg"])
            elif cls == "Canvas":
                w.configure(background=C["cell_bg"])
            elif cls == "Text":
                w.configure(background=C["field_bg"], foreground=C["fg"], insertbackground=C["fg"])
        except tk.TclError:
            pass


def _descendants(w):
    out = []
    for child in w.winfo_children():
        out.append(child)
        out += _descendants(child)
    return out


def style_menu(menu):
    """Colours for a menu created after the switch (right-click menus are built on demand)."""
    _recolor(menu)
    return menu


def _title_bar(app, dark):
    """Windows 10/11: dark title bar to match. Done when the window is idle (asking for the window handle
    right away would force a full layout and slow down opening the window)."""
    if sys.platform != "win32" or (not dark and not getattr(app, "_dark_title", False)):
        return                                   # light is the default: nothing to change
    app._dark_title = dark
    app.after_idle(lambda: _title_bar_now(app, dark))


def _title_bar_now(app, dark):
    try:
        import ctypes
        hwnd = ctypes.windll.user32.GetParent(app.winfo_id())
        value = ctypes.c_int(1 if dark else 0)
        for attr in (20, 19):   # DWMWA_USE_IMMERSIVE_DARK_MODE (newer / older builds)
            if ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, attr, ctypes.byref(value), ctypes.sizeof(value)) == 0:
                break
    except Exception:
        pass
