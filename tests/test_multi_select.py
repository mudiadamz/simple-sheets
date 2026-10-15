"""Several blocks selected with Ctrl+click: shading, formatting, Delete, status sum."""


def press(app, r, c, seq="<Button-1>"):
    g = app.table
    x, y = g.col_x(c) + g.colw(c) // 2, g.row_y(r) + g.rh(r) // 2   # centre: away from the fill handle
    g.cv.event_generate("<ButtonPress-2>", x=x, y=y, when="now")    # quick clicks would read as a double-click
    g.cv.event_generate("<ButtonRelease-2>", x=x, y=y, when="now")
    g.cv.event_generate(seq, x=x, y=y, when="now")
    g.cv.event_generate("<ButtonRelease-1>", x=x, y=y, when="now")
    app.update()


def setup(app):
    app.set_cells([(r, c, str((r + 1) * (c + 1))) for r in range(5) for c in range(4)])
    app.update()
    return app.table


def test_ctrl_click_adds_blocks(app):
    g = setup(app)
    press(app, 0, 0)
    press(app, 1, 1, "<Shift-Button-1>")              # A1:B2
    press(app, 3, 3, "<Control-Button-1>")            # + D4
    assert g.rects() == [(0, 0, 1, 1), (3, 3, 3, 3)] and g.multi()
    assert g.in_selection(1, 0) and g.in_selection(3, 3) and not g.in_selection(2, 2)
    assert app.status["text"].startswith("Count: 5") and "Sum: 25" in app.status["text"]   # 1+2+2+4 + 16
    press(app, 4, 0, "<Control-Button-1>")            # + A5
    assert len(g.rects()) == 3
    press(app, 2, 2)                                   # a plain click starts over
    assert g.rects() == [(2, 2, 2, 2)] and not g.multi()


def test_format_and_delete_all_blocks(app):
    g = setup(app)
    press(app, 0, 0)
    press(app, 2, 2, "<Control-Button-1>")
    app.format_cells(lambda st: {**st, "bold": True})
    assert g.style(0, 0) == {"bold": True} and g.style(2, 2) == {"bold": True} and not g.style(1, 1)
    app.undo()                                         # one undo step for both blocks
    assert not g.style(0, 0) and not g.style(2, 2)
    press(app, 0, 0)                                   # (undo selected what it changed)
    press(app, 2, 2, "<Control-Button-1>")
    g.clear()
    assert g.text(0, 0) == "" and g.text(2, 2) == "" and g.text(1, 1) == "4"
    app.undo()
    assert g.text(0, 0) == "1" and g.text(2, 2) == "9"


def test_keyboard_move_drops_extra_blocks(app):
    g = setup(app)
    press(app, 0, 0)
    press(app, 2, 2, "<Control-Button-1>")
    g.move(1, 0)
    assert g.extra == [] and g.cur == (3, 2)
