"""Light / dark mode switch."""
from simplecore import settings, theme


def canvas_rects(g, tag=None):
    return [g.cv.itemcget(i, "fill") for i in g.cv.find_all() if g.cv.type(i) == "rectangle"]


def test_starts_light(app):
    assert not app.dark_var.get() and theme.C["dark"] is False
    assert app.table.cv.cget("background") == theme.LIGHT["cell_bg"]


def test_switch_to_dark_and_back(app):
    g = app.table
    app.set_cells([(0, 0, "hello")])
    app.dark_var.set(True)
    app.toggle_dark()
    assert theme.C["dark"] and app.cget("background") == theme.DARK["bg"]
    assert g.cv.cget("background") == theme.DARK["cell_bg"]
    assert theme.DARK["header_bg"] in canvas_rects(g)               # headers redrawn dark
    texts = [g.cv.itemcget(i, "fill") for i in g.cv.find_withtag("txt")]
    assert texts == [theme.DARK["cell_fg"]]                          # default text turns light
    assert str(app.app_menu.cget("background")) == theme.DARK["menu_bg"]
    assert settings.get("dark") is True                              # remembered
    app.dark_var.set(False)
    app.toggle_dark()
    assert g.cv.cget("background") == theme.LIGHT["cell_bg"] and settings.get("dark") is False
    assert [g.cv.itemcget(i, "fill") for i in g.cv.find_withtag("txt")] == ["#000000"]


def test_icons_redrawn_in_light_ink(app):
    icon = app.fmtbar.icons["left"]
    dark_ink = tuple(int(theme.LIGHT["icon"][i:i + 2], 16) for i in (1, 3, 5))
    light_ink = tuple(int(theme.DARK["icon"][i:i + 2], 16) for i in (1, 3, 5))
    s = app.fmtbar.px(2)
    assert icon.get(s, s) == dark_ink
    app.set_theme(True)
    assert app.fmtbar.icons["left"] is icon and icon.get(s, s) == light_ink   # same image, new ink


def test_cell_colours_in_dark(app):
    app.set_theme(True)
    assert theme.text_color(None, None) == theme.DARK["cell_fg"]           # plain cell: light text
    assert theme.text_color("#000000", None) == theme.DARK["cell_fg"]      # black would vanish
    assert theme.text_color("#ff0000", None) == "#ff0000"                  # visible colours stay
    assert theme.text_color("#333333", None) == theme.DARK["cell_fg"]      # very dark grey too
    assert theme.text_color(None, "#ffff00") == "#000000"                  # filled cell: dark text
    app.set_theme(False)
    assert theme.text_color("#000000", None) == "#000000"


def test_next_start_remembers_dark(monkeypatch, tmp_path):
    """Dark chosen earlier -> the app opens dark (no other window open at the same time)."""
    from conftest import make_app
    monkeypatch.setenv("SIMPLE_SHEETS_HOME", str(tmp_path))
    settings.set("dark", True)
    a = make_app()
    try:
        a.attributes("-alpha", 0.0)
        assert a.dark_var.get() and theme.C["dark"]
        assert a.table.cv.cget("background") == theme.DARK["cell_bg"]
    finally:
        a.set_theme(False)
        a.destroy()


def test_next_start_dark_title_bar_when_shown(monkeypatch, tmp_path):
    """Opening already dark: the title bar is set when the window appears (before that there's no frame)."""
    import sys

    import pytest
    if sys.platform != "win32":
        pytest.skip("Windows title bar")
    from conftest import make_app
    monkeypatch.setenv("SIMPLE_SHEETS_HOME", str(tmp_path))
    settings.set("dark", True)
    done = []
    real = theme._title_bar_now
    monkeypatch.setattr(theme, "_title_bar_now", lambda w, dark: done.append((w, dark, real(w, dark))))
    a = make_app()
    try:
        a.attributes("-alpha", 0.0)
        a.update()
        assert (a, True, True) in done              # applied for real once the window was there
    finally:
        a.set_theme(False)
        a.destroy()


def test_new_windows_follow_dark(app, monkeypatch):
    import tkinter as tk
    calls = []
    monkeypatch.setattr(theme, "_title_bar_now", lambda w, dark: calls.append((w, dark)))
    app.set_theme(True)
    top = tk.Toplevel(app)
    lb, plain = tk.Listbox(top), tk.Label(top, text="hello")
    chosen = tk.Label(top, text="swatch", background="#ff0000")
    white = tk.Label(top, text="white swatch", background="#ffffff")
    for w in (lb, plain, chosen, white):
        w.pack()
    app.update()
    assert (top, True) in calls                       # dark title bar for the new window
    assert str(lb.cget("background")) == theme.DARK["field_bg"] and str(lb.cget("foreground")) == theme.DARK["fg"]
    assert str(plain.cget("background")) == theme.DARK["bg"] and str(top.cget("background")) == theme.DARK["bg"]
    assert str(chosen.cget("background")) == "#ff0000" and str(white.cget("background")) == "#ffffff"
    app.set_theme(False)                              # and back: the dialog follows
    app.update()
    assert str(lb.cget("background")) == theme.LIGHT["field_bg"] and str(plain.cget("background")) == theme.LIGHT["bg"]
    assert (top, False) in calls
    top.destroy()


def test_palette_and_contrast(app, monkeypatch):
    import palette
    monkeypatch.setattr(theme, "_title_bar_now", lambda w, dark: None)
    app.contrast_var.set(True)
    app.set_theme(False)                              # high contrast is dark even with dark mode off
    p = palette.open_palette(app)
    app.update()
    assert str(p.list.cget("background")) == theme.HIGH["field_bg"] and str(p.cget("background")) == theme.HIGH["bg"]
    p.destroy()
    app.contrast_var.set(False)
    app.set_theme(False)


def test_title_bar_needs_a_shown_window(app):
    import sys
    import tkinter as tk

    import pytest
    if sys.platform != "win32":
        pytest.skip("Windows title bar")
    top = tk.Toplevel(app)
    try:
        top.withdraw()
        assert theme._title_bar_now(top, True) is False   # not shown yet: nothing to change
        top.attributes("-alpha", 0.0)
        top.deiconify()
        top.update()
        assert theme._title_bar_now(top, True) is True
    finally:
        top.destroy()


def test_light_fills_shown_dark(app):
    def lightness(c):
        import colorsys
        return colorsys.rgb_to_hls(*(int(c[i:i + 2], 16) / 255 for i in (1, 3, 5)))[1]
    app.set_theme(True)
    try:
        band = theme.display_fill("#dde5f3")                  # pale table banding
        assert lightness(band) < 0.3 and band != "#dde5f3"
        assert theme.display_fill("#ffffff") != "#ffffff" and lightness(theme.display_fill("#ffffff")) < 0.2
        assert theme.display_fill("#0d9488") == "#0d9488"     # strong colours stay as they are
        assert theme.display_fill(None) is None
        assert theme.text_color(None, "#dde5f3") == theme.DARK["cell_fg"]       # default text: light
        assert lightness(theme.text_color("#1d4ed8", "#dbeafe")) >= 0.75         # blue on pale blue: light blue
        assert theme.text_color("#ffffff", "#0d9488") == "#ffffff"              # white on teal: unchanged
        assert theme.text_color(None, "#0d9488") == "#000000"
    finally:
        app.set_theme(False)
    assert theme.display_fill("#dde5f3") == "#dde5f3"         # light mode: real colours
    assert theme.text_color(None, "#dde5f3") == "#000000"


def test_table_banding_drawn_dark(app, tmp_path):
    import os
    import sys
    sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))
    import make_showcase

    from conftest import open_in
    g = open_in(app, make_showcase.build(str(tmp_path / "s.xlsx")))
    app.show("Orders")
    g = app.table
    app.set_theme(True)
    try:
        app.update()
        g.redraw()
        fills = {g.cv.itemcget(i, "fill").lower() for i in g.cv.find_all() if g.cv.type(i) == "rectangle"}
        ui = {v.lower() for v in theme.DARK.values() if isinstance(v, str)}       # cursor, handles…
        light = [f for f in fills - ui if f.startswith("#") and len(f) == 7 and theme.is_light_fill(f)]
        assert not light, light                                 # no pale stripes left on the dark grid
        texts = [g.cv.itemcget(i, "fill") for i in g.cv.find_all() if g.cv.type(i) == "text"]
        assert theme.DARK["cell_fg"] in texts
    finally:
        app.set_theme(False)


def test_editor_and_menus_follow_theme(app):
    g = app.table
    app.set_theme(True)
    g.goto(0, 0)
    g.edit("x")
    app.update()
    assert str(g.editor.cget("background")) == theme.DARK["field_bg"]
    g.commit(cancel=True)
    from test_menu import build_menu
    m = build_menu(app)
    assert str(m.cget("background")) == theme.DARK["menu_bg"]
    app.set_theme(False)
