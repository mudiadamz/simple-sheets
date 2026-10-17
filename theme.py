"""Light / dark colours for the whole app. `C` always holds the active palette."""
import colorsys
import functools
import os
import subprocess
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
    "chat_bg": "#f4f6f8", "chat_user": "#dce8fc", "chat_code": "#e6e9ee", "on_accent": "#ffffff",
    "danger": "#d93025",
}
DARK = {
    "dark": True,
    "bg": "#202124", "fg": "#e8eaed", "muted": "#9aa0a6",
    "cell_bg": "#1e1f22", "cell_fg": "#e8eaed", "grid_line": "#3a3b3e",
    "header_bg": "#2b2c2f", "header_sel": "#3c4a63", "header_fg": "#bdc1c6", "header_line": "#45464a",
    "accent": "#8ab4f8", "hidden_mark": "#9aa0a6", "icon": "#e8eaed",
    "field_bg": "#2b2c2f", "select_bg": "#3c4a63", "menu_bg": "#2b2c2f", "menu_active": "#3c4a63",
    "tooltip_bg": "#3c4043", "tooltip_fg": "#e8eaed",
    "chat_bg": "#26282c", "chat_user": "#2d3b52", "chat_code": "#34373c", "on_accent": "#202124",
    "danger": "#f28b82",
}
HIGH = {   # high contrast: white on black, yellow accents, strong lines
    "dark": True, "contrast": True,
    "bg": "#000000", "fg": "#ffffff", "muted": "#d0d0d0",
    "cell_bg": "#000000", "cell_fg": "#ffffff", "grid_line": "#9a9a9a",
    "header_bg": "#000000", "header_sel": "#00309c", "header_fg": "#ffffff", "header_line": "#ffffff",
    "accent": "#ffff00", "hidden_mark": "#ffff00", "icon": "#ffffff",
    "field_bg": "#000000", "select_bg": "#00309c", "menu_bg": "#000000", "menu_active": "#00309c",
    "tooltip_bg": "#000000", "tooltip_fg": "#ffff00",
    "chat_bg": "#000000", "chat_user": "#00265e", "chat_code": "#1c1c1c", "on_accent": "#000000",
    "danger": "#ff5555",
}
C = dict(LIGHT)
LIGHT_THEME = None   # the platform's native ttk theme, remembered on first use


LIGHT_FILL = 0.72    # fills lighter than this are shown as a dark shade in dark mode


@functools.lru_cache(maxsize=4096)
def _hls(hex_color):
    return colorsys.rgb_to_hls(*(int(hex_color[i:i + 2], 16) / 255 for i in (1, 3, 5)))


def _hex(h, light, s):
    return "#%02x%02x%02x" % tuple(round(x * 255) for x in colorsys.hls_to_rgb(h, light, s))


def _is_hex(c):
    return isinstance(c, str) and len(c) == 7 and c.startswith("#")


def is_light_fill(bg):
    return _is_hex(bg) and _hls(bg.lower())[1] > LIGHT_FILL


@functools.lru_cache(maxsize=4096)
def _dark_fill(bg):
    h, light, s = _hls(bg)
    return _hex(h, 0.13 + (1 - light) * 0.6, min(s, 0.55))   # white -> deep grey, pale tint -> deep tint


@functools.lru_cache(maxsize=4096)
def _light_text(fg):
    h, light, s = _hls(fg)
    return fg if light >= 0.6 else _hex(h, 0.8, min(s, 0.85))


def shade(hex_color, k):
    """A darker (k > 0) or lighter (k < 0) version of a colour."""
    h, light, s = _hls(hex_color.lower())
    return _hex(h, max(0.0, min(1.0, light * (1 - k) if k > 0 else light + (1 - light) * -k)), s)


def display_fill(bg):
    """How a cell fill is shown: as it is, except light fills in dark / high-contrast mode (table banding,
    pale table styles, pastel cards) become a dark shade of the same colour. Display only: files, PDF and
    printing keep the real colour."""
    if not bg or not C["dark"] or not is_light_fill(bg):
        return bg
    return _dark_fill(bg.lower())


def text_color(fg, bg):
    """Colour for cell text: the cell's own colour, unless it would vanish. bg is the cell's real fill."""
    if bg and not (C["dark"] and is_light_fill(bg)):   # a fill shown as it is: keep its light-mode contrast
        return fg or "#000000"
    if not fg:
        return C["cell_fg"]
    if not C["dark"]:
        return fg
    if bg:                                   # a light fill shown dark: dark text becomes a light shade of itself
        return C["cell_fg"] if near_black(fg) else (_light_text(fg.lower()) if _is_hex(fg) else fg)
    return C["cell_fg"] if near_black(fg) else fg


def near_black(hex_color):
    """Black / very dark grey: invisible on a dark cell. Colours like red or navy stay."""
    return max(int(hex_color[i:i + 2], 16) for i in (1, 3, 5)) < 0x50


def base_styles(style, scale):
    """App-specific ttk styles; re-applied after every theme switch (they are per-theme)."""
    style.configure("Bottom.TNotebook", tabposition="sw")
    if style.theme_use() == "aqua":   # macOS light: no padding (18 8 18 17) / empty sunken page above the tabs
        style.configure("Bottom.TNotebook", padding=0, tabmargins=0)
        style.layout("Bottom.TNotebook", [])
    style.configure("Bottom.TNotebook.Tab", padding=(int(14 * scale), int(3 * scale)))
    style.configure("Icon.Toolbutton", font=("Segoe UI Symbol", 12))
    for name, font in (("B", ("Segoe UI", 10, "bold")), ("I", ("Segoe UI", 10, "italic")),
                       ("U", ("Segoe UI", 10, "underline")), ("S", ("Segoe UI", 10, "overstrike")),
                       ("N", ("Segoe UI", 10))):
        style.configure(f"{name}.Toolbutton", font=font)


FOLLOW = False   # the app follows the system's light / dark (no choice of its own): macOS widgets "auto"


def system_dark(platform=None):
    """The system's light / dark setting: True (dark), False (light), None (unknown: Linux, an error).
    SIMPLE_SHEETS_SYSTEM_THEME=light / dark overrides it (tests)."""
    forced = os.environ.get("SIMPLE_SHEETS_SYSTEM_THEME", "").lower()
    if forced in ("light", "dark"):
        return forced == "dark"
    platform = platform or sys.platform
    if platform == "darwin":   # "Dark" when dark; the key is missing when light
        try:
            out = subprocess.run(["defaults", "read", "-g", "AppleInterfaceStyle"], capture_output=True,
                                 text=True, timeout=3)
        except (OSError, subprocess.SubprocessError):
            return None
        return out.stdout.strip().lower() == "dark"
    if platform == "win32":
        try:
            import winreg
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize") as key:
                return winreg.QueryValueEx(key, "AppsUseLightTheme")[0] == 0
        except (OSError, ImportError):
            return None
    return None


def _mac_appearance(win, dark):
    """macOS: native widgets / title bar in the app's light or dark ("auto" while following the system)."""
    if sys.platform != "darwin":
        return
    names = [("auto", "auto")] if FOLLOW else [("dark", "light"), ("darkaqua", "aqua")]   # Tk 9 / Tk 8.6 names
    for dark_name, light_name in names:
        try:
            win.wm_attributes("-appearance", dark_name if dark else light_name)
            return
        except tk.TclError:   # the other Tk's names, or an old Tk without -appearance
            pass


def apply(app, dark, contrast=False, follow=False):
    """Switch the whole window between light, dark and high contrast. follow: the choice is the system's."""
    global LIGHT_THEME, FOLLOW
    FOLLOW = follow and not contrast
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
    _mac_appearance(app, dark)
    for w in _descendants(app):                  # dialogs already open: their title bars too
        if w.winfo_class() == "Toplevel" and w.winfo_ismapped():
            _title_bar(w, dark)
            _mac_appearance(w, dark)
    install(app)


# window / plain-widget backgrounds that are "unset" (not a colour someone chose for that widget)
_PLAIN_BG = {"systembuttonface", "#f0f0f0", "#d9d9d9"}
_PLAIN_FG = {"systembuttontext", "systemwindowtext", "black", "#000000"}


def _plain(value, extra, keys):
    """Is this colour a default / one of our palettes' own (so it should follow the theme)?"""
    v = str(value).lower()
    return v in extra or v in {p[k].lower() for p in (LIGHT, DARK, HIGH) for k in keys}


def _recolor(widget):
    """Plain tk widgets (menus, canvas, text, lists, labels, frames in dialogs) don't follow ttk themes: set them
    directly. Labels / frames keep a colour that was chosen for them (colour swatches, tooltips)."""
    for w in [widget] + _descendants(widget):
        if getattr(w, "_own_colors", False):        # widgets that colour themselves (the AI chat)
            continue
        cls = w.winfo_class()
        try:
            if cls == "Menu":
                w.configure(background=C["menu_bg"], foreground=C["fg"], activebackground=C["menu_active"],
                            activeforeground=C["fg"], disabledforeground=C["muted"], selectcolor=C["fg"])
            elif cls == "Canvas":
                if _plain(w.cget("background"), _PLAIN_BG | {"#ffffff", "white"}, ("bg", "cell_bg")):
                    w.configure(background=C["cell_bg"])
            elif cls == "Text":
                w.configure(background=C["field_bg"], foreground=C["fg"], insertbackground=C["fg"])
            elif cls == "Listbox":
                w.configure(background=C["field_bg"], foreground=C["fg"], selectbackground=C["select_bg"],
                            selectforeground=C["fg"], highlightbackground=C["header_line"], highlightcolor=C["accent"])
            elif cls in ("Entry", "Spinbox"):
                w.configure(background=C["field_bg"], foreground=C["fg"], insertbackground=C["fg"])
            elif cls in ("Toplevel", "Frame", "Label", "Checkbutton", "Radiobutton", "Button"):
                if _plain(w.cget("background"), _PLAIN_BG, ("bg",)):
                    w.configure(background=C["bg"])
                    if cls != "Toplevel" and cls != "Frame" and _plain(w.cget("foreground"), _PLAIN_FG, ("fg",)):
                        w.configure(foreground=C["fg"])
                    if cls in ("Checkbutton", "Radiobutton", "Button"):
                        w.configure(activebackground=C["menu_active"], activeforeground=C["fg"])
                    if cls in ("Checkbutton", "Radiobutton"):
                        w.configure(selectcolor=C["field_bg"])
        except tk.TclError:
            pass


def install(app):
    """Every window, when it appears: dark / high-contrast title bar and the plain widgets recoloured
    (dialogs, the command palette, pop-ups made after the switch). Bound once."""
    if getattr(app, "_theme_hooked", False):
        return
    app._theme_hooked = True

    def on_map(e):
        w = e.widget
        if not isinstance(w, (tk.Tk, tk.Toplevel)):
            return
        _mac_appearance(w, C["dark"])
        if C["dark"] or getattr(w, "_dark_title", False):
            _recolor(w)
            _title_bar(w, C["dark"], force=True)
    app.bind_class("Toplevel", "<Map>", on_map, add="+")
    app.bind("<Map>", on_map, add="+")           # the main window (fires for its children too; filtered)


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


def _title_bar(win, dark, force=False):
    """Windows 10/11: dark title bar to match. Done when the window is idle (asking for the window handle
    right away would force a full layout and slow down opening the window). Before the window is shown
    there is no frame to change yet: the <Map> hook (install) applies it again when it appears."""
    if sys.platform != "win32" or (not dark and not getattr(win, "_dark_title", False) and not force):
        return                                   # light is the default: nothing to change
    win._dark_title = dark
    win._root().after_idle(lambda: _title_bar_now(win, dark))   # on the main window: win may close first


def _title_bar_now(win, dark):
    """-> True when the frame was changed (False: not shown yet / not Windows)."""
    try:
        import ctypes
        if not win.winfo_exists():
            return False
        user32 = ctypes.windll.user32
        hwnd = user32.GetParent(win.winfo_id())
        if not hwnd:
            return False
        value = ctypes.c_int(1 if dark else 0)
        for attr in (20, 19):   # DWMWA_USE_IMMERSIVE_DARK_MODE (newer / older builds)
            if ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, attr, ctypes.byref(value), ctypes.sizeof(value)) == 0:
                break
        # repaint the frame now (else the old colour stays until the window is moved or refocused)
        user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0, 0x0001 | 0x0002 | 0x0004 | 0x0010 | 0x0020)
        return True
    except Exception:
        return False
