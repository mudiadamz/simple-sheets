"""Name box / Go to: type a reference to jump there."""
from conftest import key
from grid import ALL


def type_ref(app, text):
    key(app, "<Control-g>")
    assert app.focus_get() is app.ref
    app.ref.delete(0, "end")
    app.ref.insert(0, text)
    key(app, "<KeyPress-x>", widget=app.ref)              # a real key press marks "typing"
    app.ref.delete(0, "end")
    app.ref.insert(0, text)
    assert app.ref.typing
    key(app, "<Return>", widget=app.ref)
    assert app.ref.get() == app.ref["text"] and not app.ref.typing


def test_shows_selection(app):
    app.table.select(1, 0, 2, 2)
    assert app.ref.get() == "A2:C3"


def test_go_to_cell_range_rows_columns(app):
    g = app.table
    type_ref(app, "d12")
    assert g.cur == (11, 3) and not g.multi() and app.focus_get() is g.cv
    assert app.ref.get() == "D12"                         # box shows the new position
    type_ref(app, "B2:C5")
    assert g.sel() == (1, 1, 4, 2)
    type_ref(app, "C:D")
    assert g.sel() == (0, 2, ALL, 3)
    type_ref(app, "3:4")
    assert g.sel() == (2, 0, 3, ALL)


def test_go_to_other_sheet(app):
    app.add_sheet()
    app.rename_sheet("Sheet2", "My Data")
    app.show("Sheet1")
    type_ref(app, "'My Data'!B3")
    assert app.sheet == "My Data" and app.table.cur == (2, 1)


def test_bad_reference_warns(app):
    type_ref(app, "not a ref")
    assert any(d[0] == "showwarning" for d in app.dialogs)
    assert app.ref.get() == "A1"                         # back to showing the selection
