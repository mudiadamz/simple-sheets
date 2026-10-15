"""Paste special: transpose, add / subtract / multiply / divide, skip blanks."""


def copy_block(app, r1, c1, r2, c2):
    g = app.table
    g.select(r1, c1, r2, c2)
    g.copy()


def test_transpose_keeps_formulas_right(app):
    g = app.table
    app.set_cells([(0, 0, "1"), (0, 1, "2"), (0, 2, "=A1+B1")])
    copy_block(app, 0, 0, 0, 2)
    g.goto(2, 0)
    g.paste(transpose=True)
    assert [g.text(r, 0) for r in (2, 3, 4)] == ["1", "2", "=A3+A4"]
    assert g.value(4, 0) == "3"


def test_operations(app):
    g = app.table
    app.set_cells([(0, 0, "10"), (1, 0, "20"), (0, 1, "2"), (1, 1, "x"), (1, 3, "=5")])
    copy_block(app, 0, 1, 1, 1)                       # 2 and "x"
    g.goto(0, 0)
    g.paste(op="multiply")
    assert g.text(0, 0) == "20" and g.text(1, 0) == "20"      # "x" isn't a number: left alone
    app.set_cells([(5, 0, "3")])
    copy_block(app, 5, 0, 5, 0)
    g.goto(1, 3)
    g.paste(op="add")
    assert g.text(1, 3) == "=(5)+3" and g.value(1, 3) == "8"  # a formula stays a formula
    g.goto(6, 0)
    g.paste(op="divide")
    assert g.text(6, 0) == "0"                         # empty target counts as 0
    app.set_cells([(7, 0, "0")])
    copy_block(app, 7, 0, 7, 0)
    g.goto(0, 0)
    g.paste(op="divide")
    assert g.text(0, 0) == "#DIV/0!"


def test_skip_blanks(app):
    g = app.table
    app.set_cells([(0, 0, "a"), (2, 0, "c"), (0, 2, "keep1"), (1, 2, "keep2"), (2, 2, "keep3")])
    copy_block(app, 0, 0, 2, 0)
    g.goto(0, 2)
    g.paste(skip_blanks=True)
    assert [g.text(r, 2) for r in range(3)] == ["a", "keep2", "c"]
    app.undo()
    assert g.text(0, 2) == "keep1"
