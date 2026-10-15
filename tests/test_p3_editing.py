"""Editing helpers: format painter, cell styles, column autocomplete, drag a selection to move it,
arrow keys pointing at cells while typing a formula."""
from conftest import key


def test_format_painter(app):
    g = app.table
    app.set_cells([(0, 0, "a"), (0, 1, "b"), (2, 0, "x"), (2, 1, "y"), (2, 2, "z")])
    g.goto(0, 0)
    app.format_cells(lambda st: {**st, "bold": True, "bg": "#ffff00"})
    app.format_painter()
    assert app.painter == [[{"bold": True, "bg": "#ffff00"}]]
    g.select(2, 0, 2, 2)
    g.drag = "cells"
    g.on_release(type("E", (), {"x": 0, "y": 0, "state": 0})())
    assert all(g.style(2, c) == {"bold": True, "bg": "#ffff00"} for c in range(3))
    assert app.painter is None
    app.undo()
    assert not g.style(2, 1)
    app.format_painter()
    app.cancel_painter()
    assert app.painter is None


def test_cell_styles(app):
    g = app.table
    app.set_cells([(0, 0, "ok")])
    app.format_cells(lambda st: {**st, "numfmt": "0.00"})
    app.fmtbar.cell_style("Good")
    assert g.style(0, 0) == {"numfmt": "0.00", "bg": "#c6efce", "fg": "#006100"}
    app.fmtbar.cell_style("Heading 1")
    assert g.style(0, 0)["border"]["bottom"][0] == "thick" and "bg" not in g.style(0, 0)
    app.fmtbar.cell_style("Normal")
    assert g.style(0, 0) == {"numfmt": "0.00"}
    labels = [app.fmtbar.styles_menu.entrycget(i, "label") for i in range(app.fmtbar.styles_menu.index("end") + 1)]
    assert {"Good", "Bad", "Neutral", "Title", "Total", "Heading 4"} <= set(labels)


def test_column_autocomplete(app):
    g = app.table
    app.set_cells([(0, 0, "Apple"), (1, 0, "Apricot"), (2, 0, "Banana")])
    g.goto(3, 0)
    g.edit("B")
    e = g.editor
    g.column_complete(e)
    assert e.get() == "Banana" and e.get(*[int(x) for x in ("1", "6")]) == "anana"
    g.commit()
    assert g.text(3, 0) == "Banana"
    g.goto(4, 0)
    g.edit("Ap")
    g.column_complete(g.editor)
    assert g.editor.get() == "Ap"                 # two candidates: no suggestion
    g.editor.insert("end", "r")
    g.column_complete(g.editor)
    assert g.editor.get() == "Apricot"
    g.commit(cancel=True)


def test_move_block(app):
    g = app.table
    app.set_cells([(0, 0, "1"), (0, 1, "2"), (1, 0, "3"), (5, 5, "keep")])
    app.format_cells(lambda st: st)
    app.set_comment(0, 0, "note")
    assert app.move_block((0, 0, 1, 1), 3, 2)
    assert g.text(0, 0) == "" and [g.text(3, 2), g.text(3, 3), g.text(4, 2)] == ["1", "2", "3"]
    assert (3, 2) in g.sheet.comments and (0, 0) not in g.sheet.comments
    assert g.sel() == (3, 2, 4, 3)
    app.undo()
    assert g.text(0, 0) == "1" and g.text(3, 2) == ""


def test_drag_edge_moves(app):
    g = app.table
    app.set_cells([(0, 0, "v")])
    app.update()
    g.goto(0, 0)
    app.update()
    x, y = g.col_x(1) - 1, g.row_y(0) + g.rh(0) // 2      # right edge of A1
    assert g.on_sel_edge(x, y) and not g.on_sel_edge(x - 15, y)
    g.cv.event_generate("<ButtonPress-1>", x=x, y=y, when="now")
    tx, ty = g.col_x(2) + 10, g.row_y(2) + g.rh(2) // 2
    g.cv.event_generate("<B1-Motion>", x=tx, y=ty, when="now")
    assert g.fill_preview is not None
    g.cv.event_generate("<ButtonRelease-1>", x=tx, y=ty, when="now")
    app.update()
    assert g.text(0, 0) == "" and g.text(2, 2) == "v"


def test_arrow_keys_point_in_formula(app):
    g = app.table
    app.set_cells([(r, 0, str(r + 1)) for r in range(5)])
    g.goto(5, 0)
    g.edit("=")
    e = g.editor
    key(app, "<Up>", e)
    assert e.get() == "=A5"
    key(app, "<Up>", e)
    assert e.get() == "=A4"                       # moves the same reference
    key(app, "<Shift-Up>", e)
    key(app, "<Shift-Up>", e)
    assert e.get() == "=A2:A4"
    e.insert("end", "+")
    e.icursor("end")
    key(app, "<Right>", e)
    assert e.get() == "=A2:A4+B6"                 # a new reference starts at the edited cell
    g.commit()
    assert g.text(5, 0) == "=A2:A4+B6"


def test_arrows_still_finish_typing(app):
    g = app.table
    g.goto(0, 0)
    g.edit("x")
    key(app, "<Down>", g.editor)
    assert g.text(0, 0) == "x" and g.cur == (1, 0)
