"""Dashed colored boxes around the cells used by a formula while it is being typed."""
import pytest

import workbooks
from conftest import key, open_in
from formula import references


@pytest.mark.parametrize("text, want", [
    ("=A1+B2", [(None, 0, 0, 0, 0), (None, 1, 1, 1, 1)]),
    ("=SUM(B2:C4)*2", [(None, 1, 1, 3, 2)]),
    ("=SUM(A:A)", [(None, 0, 0, 1048575, 0)]),
    ("='My Sheet'!$B$3", [("My Sheet", 2, 1, 2, 1)]),
    ('="A1"&C1', [(None, 0, 2, 0, 2)]),     # text in quotes is not a reference
    ("=SUM(A1:", [(None, 0, 0, 0, 0)]),       # half-typed formula
    ("A1+B1", []),                           # not a formula
])
def test_references(text, want):
    assert references(text) == want


def dashed_boxes(g):
    return [g.cv.itemcget(i, "outline") for i in g.cv.find_withtag("ref")]


def type_into(app, entry, text):
    entry.delete(0, "end")
    entry.insert(0, text)
    key(app, "<KeyRelease>", widget=entry)


def test_highlight_while_editing_formula(app, tmp_path):
    g = open_in(app, workbooks.budget(tmp_path / "b.xlsx"))
    g.goto(7, 0)
    g.edit("=")
    app.update()
    assert g.ref_marks == [] and dashed_boxes(g) == []
    type_into(app, g.editor, "=B2*C2+B2")
    assert [m[:4] for m in g.ref_marks] == [(1, 1, 1, 1), (1, 2, 1, 2), (1, 1, 1, 1)]
    colors = [m[4] for m in g.ref_marks]
    assert colors[0] == colors[2] != colors[1]          # same cell -> same color
    assert len(dashed_boxes(g)) == 3
    type_into(app, g.editor, "=SUM(D2:D4)")
    assert [m[:4] for m in g.ref_marks] == [(1, 3, 3, 3)]
    narrow = g.editor.winfo_width()
    type_into(app, g.editor, "=SUM(D2:D4)+SUM(B2:B4)+SUM(C2:C4)")
    assert g.editor.winfo_width() > narrow           # editor grows while typing
    g.commit()
    assert g.ref_marks == [] and dashed_boxes(g) == []


def test_editing_existing_formula_shows_refs(app, tmp_path):
    g = open_in(app, workbooks.budget(tmp_path / "b.xlsx"))
    g.goto(5, 3)            # =SUM(D2:D4)
    g.edit()
    app.update()
    assert [m[:4] for m in g.ref_marks] == [(1, 3, 3, 3)]
    g.commit(cancel=True)
    assert g.ref_marks == []


def test_plain_text_and_other_sheet_refs_not_highlighted(app, tmp_path):
    g = open_in(app, workbooks.layout(tmp_path / "m.xlsx"))
    g.goto(9, 0)
    g.edit("x")
    type_into(app, g.editor, "A1 is text")
    assert g.ref_marks == []
    type_into(app, g.editor, "=Other!A1+Main!B2")
    assert [m[:4] for m in g.ref_marks] == [(1, 1, 1, 1)]
    g.commit(cancel=True)


def test_highlight_from_formula_bar(app, tmp_path):
    g = open_in(app, workbooks.budget(tmp_path / "b.xlsx"))
    g.goto(7, 0)
    app.value.focus_force()
    type_into(app, app.value, "=C3/B3")
    assert [m[:4] for m in g.ref_marks] == [(2, 2, 2, 2), (2, 1, 2, 1)]
    app.value_commit()
    app.update()
    assert g.ref_marks == [] and g.text(7, 0) == "=C3/B3"
