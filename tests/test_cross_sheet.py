"""Formulas that use other sheets, built by clicking (Excel style)."""
import workbooks
from conftest import key, open_in
from test_point_refs import click


def focus(app, w):
    for _ in range(50):
        w.focus_force()
        app.update()
        if app.focus_get() is w:
            return


def test_typed_cross_sheet_formula(app, tmp_path):
    g = open_in(app, workbooks.layout(tmp_path / "m.xlsx"))
    app.show("Other")
    app.set_cells([(2, 0, "=Main!B2+Main!C3*2")])
    assert g.value(2, 0) == "41"
    app.show("Main")
    app.set_cells([(1, 1, "5")])                           # change on Main updates Other
    app.show("Other")
    assert g.value(2, 0) == "45"


def test_point_at_other_sheet_from_cell_editor(app, tmp_path):
    g = open_in(app, workbooks.layout(tmp_path / "m.xlsx"))
    app.show("Main")
    g.goto(9, 0)
    g.edit("=")
    app.update()
    app.tabs.select("Other")                               # click the "Other" tab
    app.tabs.on_tab()
    focus(app, app.value)
    assert app.sheet == "Other" and app.formula_home[1:] == (9, 0)
    assert g.editor is None and app.value.get() == "="     # formula moved to the formula bar
    click(app, 1, 0)
    assert app.value.get() == "=Other!A2"
    assert [m[:4] for m in g.ref_marks] == [(1, 0, 1, 0)]
    app.value.insert("end", "+")
    app.value.icursor("end")
    app.show("Main")                                       # back to the home sheet
    app.tabs.select("Main")
    focus(app, app.value)
    click(app, 1, 1)
    assert app.value.get() == "=Other!A2+B2"               # home-sheet refs have no prefix
    assert [m[:4] for m in g.ref_marks] == [(1, 1, 1, 1)]  # Other!A2 isn't on this sheet
    key(app, "<Return>", widget=app.value)
    assert app.sheet == "Main" and app.formula_home is None
    assert g.text(9, 0) == "=Other!A2+B2" and g.value(9, 0) == "7"   # 6 (=SUM(Main!B2:B4)) + 1
    assert g.cur == (10, 0)


def test_escape_cancels_cross_sheet_formula(app, tmp_path):
    g = open_in(app, workbooks.layout(tmp_path / "m.xlsx"))
    app.show("Main")
    g.goto(9, 0)
    g.edit("=")
    app.update()
    app.show("Other")
    focus(app, app.value)
    click(app, 0, 0)
    key(app, "<Escape>", widget=app.value)
    assert app.sheet == "Main" and app.formula_home is None
    assert g.text(9, 0) == "" and app.hpos == 0


def test_click_that_does_not_point_finishes_formula(app, tmp_path):
    g = open_in(app, workbooks.layout(tmp_path / "m.xlsx"))
    app.show("Main")
    g.goto(9, 0)
    g.edit("=")
    app.update()
    app.show("Other")
    focus(app, app.value)
    click(app, 0, 0)                                       # =Other!A1
    app.value.insert("end", "*2")
    app.value.icursor("end")
    click(app, 3, 3)                                       # caret after "2": not a pointing place
    assert app.sheet == "Main" and g.text(9, 0) == "=Other!A1*2" and g.value(9, 0) == "280"


def test_sheet_names_with_spaces_are_quoted(app, tmp_path):
    g = open_in(app, workbooks.layout(tmp_path / "m.xlsx"))
    app.rename_sheet("Other", "My Data")
    app.show("Main")
    g.goto(9, 0)
    g.edit("=")
    app.update()
    app.show("My Data")
    focus(app, app.value)
    click(app, 1, 0)
    assert app.value.get() == "='My Data'!A2"
    app.value_commit()
    assert g.value(9, 0) == "6"
