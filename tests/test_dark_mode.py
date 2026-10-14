"""Light / dark mode switch."""
import settings
import theme


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
    assert str(app.fmt_info.cget("foreground")) == theme.DARK["muted"]
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


def test_editor_and_menus_follow_theme(app):
    import menus
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
