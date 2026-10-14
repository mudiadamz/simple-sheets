"""Hiding / unhiding whole sheets."""
from types import SimpleNamespace

import openpyxl

import sheet_io as io
import sheet_tabs
import workbooks
from conftest import open_in


def tab_menu(app, name):
    """The tab's right-click menu, built but not shown -> {label: state}, Unhide submenu labels."""
    app.update()                       # tabs laid out before measuring them
    app.tabs.menu = None
    orig = sheet_tabs.tk.Menu.tk_popup
    sheet_tabs.tk.Menu.tk_popup = lambda self, x, y: None
    try:
        x1, x2 = app.tabs.tab_box(name)
        app.tabs.menu_at(SimpleNamespace(x=(x1 + x2) // 2, y=app.tabs.nb.winfo_height() // 2,
                                         x_root=0, y_root=0))
    finally:
        sheet_tabs.tk.Menu.tk_popup = orig
    m = app.tabs.menu
    states = {m.entrycget(i, "label"): m.entrycget(i, "state")
              for i in range(m.index("end") + 1) if m.type(i) != "separator"}
    sub = m.nametowidget(m.entrycget(m.index("Unhide"), "menu"))
    subs = [sub.entrycget(i, "label") for i in range(sub.index("end") + 1)] if sub.index("end") is not None else []
    return states, subs


def test_load_hidden_sheets(app, tmp_path):
    g = open_in(app, workbooks.hidden_sheet(tmp_path / "h.xlsx"))
    assert list(app.sheets) == ["A", "Secret", "B", "Deep"]
    assert app.tabs.names() == ["A", "B"]
    app.show("B")
    assert g.value(0, 0) == "84"                           # formulas still read hidden sheets
    assert not app.dirty()


def test_hide_and_unhide_with_undo(app, tmp_path):
    open_in(app, workbooks.layout(tmp_path / "m.xlsx"))
    app.add_sheet()                                        # Main, Sheet3, Other
    assert app.hide_sheet("Sheet3")
    assert app.tabs.names() == ["Main", "Other"] and app.sheet == "Other" and app.dirty()
    assert app.hide_sheet("Other")
    assert app.tabs.names() == ["Main"] and app.sheet == "Main"
    assert not app.hide_sheet("Main")                      # the last visible sheet stays
    assert any(d[0] == "showwarning" for d in app.dialogs)
    assert not app.delete_sheet("Main")
    app.hide_sheet("Sheet3", False)
    assert app.tabs.names() == ["Main", "Sheet3"] and app.sheet == "Sheet3"
    app.undo()
    assert app.tabs.names() == ["Main"]
    app.undo()
    app.undo()
    assert app.tabs.names() == ["Main", "Sheet3", "Other"]


def test_tab_menu(app, tmp_path):
    open_in(app, workbooks.hidden_sheet(tmp_path / "h.xlsx"))
    states, subs = tab_menu(app, "A")
    assert states["Hide"] == "normal" and states["Unhide"] == "normal"
    assert subs == ["Secret", "Deep"]
    app.hide_sheet("B")
    states, _ = tab_menu(app, "A")
    assert states["Hide"] == "disabled" and states["Delete"] == "disabled"


def test_drag_order_keeps_hidden_slots(app, tmp_path):
    open_in(app, workbooks.hidden_sheet(tmp_path / "h.xlsx"))
    app.move_sheet("B", 0)                                  # visible order: B, A
    assert list(app.sheets) == ["B", "Secret", "A", "Deep"] and app.tabs.names() == ["B", "A"]


def test_save_hidden_sheets(app, tmp_path):
    path = workbooks.hidden_sheet(tmp_path / "h.xlsx")
    open_in(app, path)
    app.hide_sheet("Secret", False)                         # unhide
    app.hide_sheet("A")                                     # hide the first sheet
    assert app.save(), app.dialogs
    wb = openpyxl.load_workbook(path)
    states = {ws.title: ws.sheet_state for ws in wb.worksheets}
    assert states == {"A": "hidden", "Secret": "visible", "B": "visible", "Deep": "veryHidden"}
    assert not wb.worksheets[wb.index(wb.active)].sheet_state == "hidden"   # active sheet is visible
    again = io.load(str(path))
    assert [n for n, s in again.items() if s.hidden] == ["A", "Deep"]
    assert not app.dirty()
