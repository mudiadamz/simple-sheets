"""Moving the cursor without scrolling redraws only the selection layer, not every cell."""


def test_cursor_move_skips_cells(app):
    g = app.table
    app.set_cells([(r, c, f"{r}{c}") for r in range(10) for c in range(5)])
    app.update()
    g.goto(2, 2)
    calls = []
    real = g.draw_cell
    g.draw_cell = lambda *a, **k: (calls.append(a[:2]), real(*a, **k))[1]
    g.move(1, 0)
    g.move(0, 1)
    assert calls == [] and g.cur == (3, 3)
    boxes = [i for i in g.cv.find_withtag("ov") if g.cv.type(i) == "rectangle"]
    assert boxes                                      # the active-cell box was drawn again
    g.select(0, 0, 2, 2)                              # a block: shading needs the cells
    g.move(1, 0)
    assert calls                                      # full redraw again
    del g.draw_cell


def test_scrolling_still_redraws(app):
    g = app.table
    app.set_cells([(r, 0, str(r)) for r in range(200)])
    app.update()
    g.goto(0, 0)
    calls = []
    real = g.draw_cell
    g.draw_cell = lambda *a, **k: (calls.append(a[:2]), real(*a, **k))[1]
    g.goto(150, 0)                                    # far down: the view scrolls
    assert calls and g.top > 0
    del g.draw_cell
