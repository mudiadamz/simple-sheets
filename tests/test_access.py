"""Accessibility: high contrast, interface size, colour-blind charts, keyboard-only use, command palette, tips."""

import tkinter as tk

import charts
import keyboard_nav
import palette
import settings
import settings_ui
import theme


def test_high_contrast(app):
    app.contrast_var.set(True)
    app.toggle_contrast()
    assert theme.C["contrast"] and theme.C["cell_bg"] == "#000000" and theme.C["accent"] == "#ffff00"
    assert settings.get("contrast") is True
    assert app.table.cv["background"] == "#000000"
    app.contrast_var.set(False)
    app.toggle_contrast()
    assert not theme.C.get("contrast") and theme.C["cell_bg"] == "#ffffff"


def test_view_menu_has_high_contrast(app):
    labels = [app.view_menu.entrycget(i, "label") for i in range(app.view_menu.index("end") + 1)
              if app.view_menu.type(i) != "separator"]
    assert "High contrast" in labels and "Page breaks" in labels


def test_ui_scale_factor(app):
    from tkinter import font as tkfont
    fonts = {n: tkfont.nametofont(n, root=app) for n in settings_ui.NAMED_FONTS}
    sizes = {n: f.cget("size") for n, f in fonts.items()}     # ui_factor scales ALL of them: put them ALL back
    f, before = fonts["TkTooltipFont"], sizes["TkTooltipFont"]
    settings.set("ui_scale", 150)
    try:
        assert settings_ui.ui_factor(app) == 1.5
        assert abs(f.cget("size")) == round(abs(before) * 1.5) or before == 0
    finally:
        for n, font in fonts.items():
            font.configure(size=sizes[n])
        settings.set("ui_scale", 100)
    assert all(fonts[n].cget("size") == sizes[n] for n in fonts)
    assert settings_ui.ui_factor(app) == 1.0
    settings.set("ui_scale", "huge")
    assert settings_ui.ui_factor(app) == 1.0


def test_colour_blind_palette(app):
    settings.set("chart_palette", "colour-blind")
    try:
        settings_ui.apply_startup()
        assert charts.PALETTE[0] == "#0072b2"
    finally:
        settings.set("chart_palette", "office")
        settings_ui.apply_startup()
    assert charts.PALETTE[0] == "#4472c4"


def test_settings_dialog_saves_new_options(app):
    d = settings_ui.SettingsDialog(app)
    d.size.set("125%")
    d.palette.set(settings_ui.PALETTES["colour-blind"])
    assert d.ok()
    try:
        assert settings.get("ui_scale") == 125 and settings.get("chart_palette") == "colour-blind"
        assert charts.PALETTE == charts.COLOUR_BLIND
    finally:
        settings.set("ui_scale", 100)
        settings.set("chart_palette", "office")
        settings_ui.apply_startup()


def chart_and_picture(app):
    app.set_cells([(0, 0, "Name"), (0, 1, "Qty"), (1, 0, "b"), (1, 1, "2"), (2, 0, "c"), (2, 1, "5")])
    app.table.select(0, 0, 2, 1)
    ch = charts.insert(app, "bar", "Test")
    im = {"cell": (10, 1), "end": None, "width": 20, "height": 20, "format": "png", "data": b"",
          "photo_failed": True}
    app.current.extra["images"] = [im]
    return ch, im


def test_pick_objects_by_keyboard(app):
    ch, im = chart_and_picture(app)
    g = app.table
    assert keyboard_nav.next_object(app) == ("chart", ch) and g.selected_chart is ch
    assert keyboard_nav.tab(app, 1) == "break" and g.selected_image is im and g.selected_chart is None
    assert keyboard_nav.next_object(app, -1) == ("chart", ch)
    g.unpick()
    assert keyboard_nav.tab(app, 1) is None                         # nothing picked: Tab moves the cursor
    app.current.charts.clear()
    app.current.extra.pop("images")
    assert keyboard_nav.next_object(app) is None and "No charts" in app.status["text"]


def test_f6_cycles_panes(app):
    app.table.cv.focus_force()
    app.update()
    assert keyboard_nav.cycle_focus(app) is app.value
    app.update()
    assert keyboard_nav.cycle_focus(app) is app.ref
    app.update()
    assert keyboard_nav.cycle_focus(app) is app.tabs
    app.update()
    assert keyboard_nav.cycle_focus(app) is app.table.cv
    app.update()
    assert keyboard_nav.cycle_focus(app, -1) is app.tabs


def test_sheet_tabs_keys(app):
    app.add_sheet()
    first = app.visible_sheets()[0]
    app.show(first)
    keyboard_nav.tabs_key(app, "Right")
    assert app.sheet != first
    keyboard_nav.tabs_key(app, "Left")
    assert app.sheet == first


def test_command_palette(app):
    p = palette.open_palette(app)
    labels = [label for label, _m, _i in p.items]
    assert any(label.endswith("Insert function…") for label in labels) and any("Open…" == x for x in labels)
    p.entry.insert(0, "dark mod")
    assert p.fill() == 1 and "Dark mode" in p.list.get(0)
    p.entry.delete(0, "end")
    p.entry.insert(0, "zzzz")
    assert p.fill() == 0 and p.run() is None
    p.entry.delete(0, "end")
    p.entry.insert(0, "page breaks")
    p.fill()
    assert p.run().endswith("Page breaks") and app.page_breaks is True
    assert palette.score("ins fun", "Tools › Insert function…") == 0
    assert palette.score("fun", "Tools › Insert function…") == 0 and palette.score("x y", "abc") is None


def test_palette_runs_postcommand_given_as_tcl_object(app):
    """In the exe, cget('postcommand') returns a Tcl_Obj, not a str (eval() refused it)."""
    filled = []
    top = tk.Menu(app, tearoff=False)
    sub = tk.Menu(top, tearoff=False, postcommand=lambda: filled.append(1) or sub.add_command(label="Late item"))
    top.add_cascade(label="Sub", menu=sub)
    real_cget = tk.Menu.cget

    class TclObj:                       # stands in for _tkinter.Tcl_Obj: str() gives the script
        def __init__(self, s):
            self.s = s

        def __str__(self):
            return self.s
    try:
        tk.Menu.cget = lambda self, key: TclObj(str(real_cget(self, key))) if key == "postcommand" \
            else real_cget(self, key)
        found = [label for label, _m, _i in palette.commands(top)]
    finally:
        tk.Menu.cget = real_cget
    assert filled and ("Sub", "Late item") in found
    top.destroy()


def test_tips(app):
    settings.set("tips_at_start", True)
    t = palette.first_run(app)
    first = t.text["text"]
    t.go(1)
    assert t.text["text"] != first and t.head["text"].startswith("Tip 2")
    t.again.set(False)
    settings.set("tips_at_start", False)                          # what the check box does
    t.close()
    assert palette.first_run(app) is None and settings.get("tip_next") == 2
    w = palette.show_tips(app)
    assert w.head["text"].startswith("Tip 3")
    w.destroy()
