"""Hiding / unhiding rows and columns."""
import openpyxl

import sheet_io as io
import workbooks
from conftest import key, open_in
from test_menu import menu_labels, menu_states


def test_load_hidden_from_file(tmp_path):
    s = io.load(str(workbooks.hidden(tmp_path / "h.xlsx")))["H"]
    assert s.hidden_rows == {2, 3} and s.hidden_cols == {2}


def test_hidden_have_no_size(app, tmp_path):
    g = open_in(app, workbooks.hidden(tmp_path / "h.xlsx"))
    assert g.rh(2) == 0 and g.rh(3) == 0 and g.rh(4) > 0
    assert g.colw(2) == 0 and g.colw(3) > 0
    assert g.row_y(4) == g.row_y(2)                    # row 5 sits right under row 2
    assert g.hit(g.col_x(3) + 3, g.row_y(4) + 3) == (4, 3)
    assert len(g.cv.find_withtag("hidden")) == 2      # one marker for the rows, one for the column


def test_arrows_skip_hidden(app, tmp_path):
    g = open_in(app, workbooks.hidden(tmp_path / "h.xlsx"))
    g.goto(1, 1)
    key(app, "<Down>")
    assert g.cur == (4, 1)
    key(app, "<Up>")
    assert g.cur == (1, 1)
    key(app, "<Right>")
    assert g.cur == (1, 3)
    key(app, "<Left>")
    assert g.cur == (1, 1)


def test_hide_unhide_with_undo(app, tmp_path):
    g = open_in(app, workbooks.plain(tmp_path / "p.xlsx"))
    g.select(1, 0, 1, 2)
    app.hide("row")
    assert g.sheet.hidden_rows == {1} and g.rh(1) == 0 and app.dirty()
    assert g.cur != (1, 0)                              # cursor moved off the hidden row
    g.select(0, 1, 0, 2)
    app.hide("col")
    assert g.sheet.hidden_cols == {1, 2}
    app.undo()
    assert g.sheet.hidden_cols == set()
    app.redo()
    g.select(0, 0, 3, 3)                                # select across them, then unhide
    app.hide("row", False)
    app.hide("col", False)
    assert g.sheet.hidden_rows == set() and g.sheet.hidden_cols == set()
    assert g.sel() == (0, 0, 3, 3)                      # selection kept
    for _ in range(4):
        app.undo()
    assert not app.dirty()


def test_menu_items(app, tmp_path):
    g = open_in(app, workbooks.hidden(tmp_path / "h.xlsx"))
    g.select(0, 0, 1, 1)
    labels = menu_labels(app)
    assert "Hide rows 1 - 2" in labels and "Hide columns A - B" in labels
    assert menu_states(app)["Unhide rows"] == "disabled"
    g.select(1, 1, 4, 3)                                # spans hidden rows 3-4 and column C
    states = menu_states(app)
    assert states["Unhide rows"] == "normal" and states["Unhide columns"] == "normal"


def test_insert_row_shifts_hidden(app, tmp_path):
    g = open_in(app, workbooks.hidden(tmp_path / "h.xlsx"))
    app.restructure("row", 0, 1)
    app.restructure("col", 0, 1)
    assert g.sheet.hidden_rows == {3, 4} and g.sheet.hidden_cols == {3}
    app.restructure("row", 3, -1)
    assert g.sheet.hidden_rows == {3}
    app.undo()
    assert g.sheet.hidden_rows == {3, 4}


def test_save_hidden(app, tmp_path):
    path = workbooks.hidden(tmp_path / "h.xlsx")
    g = open_in(app, path)
    g.select(0, 0, 6, 6)
    app.hide("row", False)                              # unhide rows 3-4
    g.select(0, 5, 0, 5)
    app.hide("col")                                     # hide F, inside the E:G group
    g.select(5, 0, 5, 0)
    app.hide("row")                                     # hide row 6
    assert app.save(), app.dialogs
    ws = openpyxl.load_workbook(path)["H"]
    hidden_rows = {r for r, d in ws.row_dimensions.items() if d.hidden}
    assert hidden_rows == {6}
    again = io.load(str(path))["H"]
    assert again.hidden_cols == {2, 5} and again.hidden_rows == {5}
    assert ws.column_dimensions["E"].width == 15        # the split group kept its width


def test_save_hidden_after_insert(app, tmp_path):
    path = workbooks.hidden(tmp_path / "h.xlsx")
    open_in(app, path)
    app.restructure("row", 0, 2)                        # rebuild path
    assert app.save(), app.dialogs
    again = io.load(str(path))["H"]
    assert again.hidden_rows == {4, 5} and again.hidden_cols == {2}
