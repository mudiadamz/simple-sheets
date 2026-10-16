"""Moving the cursor without scrolling redraws only the selection layer, not every cell."""
import theme


class Wheel:
    def __init__(self, delta):
        self.delta, self.x, self.y = delta, 5, 5


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


def test_fast_wheel_gives_one_redraw(app):
    """Wheel events that arrive together scroll the whole way but draw the sheet once (no lag behind the mouse)."""
    g = app.table
    app.set_cells([(r, 0, str(r)) for r in range(300)])
    app.update()
    redraws = []
    real = g.redraw
    g.redraw = lambda: (redraws.append(g.top), real())[1]
    for _ in range(10):
        g.on_wheel(Wheel(-120))
    assert g.top == 30 and redraws == []              # moved at once, drawing waits for the events
    app.update()
    assert redraws == [30] and g.view[0][0] == 30
    g.yview("scroll", 5, "units")
    g.yview("scroll", 5, "units")
    app.update()
    assert redraws == [30, 40]
    del g.redraw


def test_touchpad_small_deltas_add_up(app):
    g = app.table
    app.set_cells([(r, 0, str(r)) for r in range(100)])
    app.update()
    for _ in range(3):
        g.on_wheel(Wheel(-20))                        # half a row each
    assert g.top == 1
    g.on_wheel(Wheel(-120))                           # a mouse notch: 3 rows
    assert g.top == 4
    g.on_wheel(Wheel(120))
    assert g.top == 1
    g.on_wheel(Wheel(-120), horizontal=True)          # Shift+wheel: one column
    assert g.left == 1


def test_gridlines_are_lines_not_a_box_per_cell(app):
    """Empty cells draw no item of their own: gridlines are one line per row / column."""
    g = app.table
    app.update()
    g.redraw()
    rows, cols = g.view
    cell_rects = [i for i in g.cv.find_all() if g.cv.type(i) == "rectangle" and "ov" not in g.cv.gettags(i)]
    assert len(cell_rects) < len(rows) * len(cols) / 4
    lines = [i for i in g.cv.find_all() if g.cv.type(i) == "line" and g.cv.itemcget(i, "fill") == theme.C["grid_line"]]
    assert len(lines) >= len(rows) + len(cols)
    app.current.props["grid"] = False                 # View → no gridlines: no grid lines drawn
    g.redraw()
    assert not [i for i in g.cv.find_all() if g.cv.itemcget(i, "fill") == theme.C["grid_line"]
                and g.cv.type(i) == "line"]
    app.current.props.pop("grid")
    g.redraw()


def test_merged_block_hides_gridlines(app):
    g = app.table
    app.set_cells([(0, 0, "title")])
    g.select(0, 0, 1, 3)
    app.merge()
    app.update()
    g.redraw()
    x1, y1 = g.col_x(0), g.row_y(0)
    x2, y2 = g.col_x(4), g.row_y(2)
    covers = [i for i in g.cv.find_all() if g.cv.type(i) == "rectangle"
              and g.cv.itemcget(i, "fill") == theme.C["cell_bg"]
              and [round(v) for v in g.cv.coords(i)] == [x1, y1, x2, y2]]
    assert covers                                     # the merged block is filled over the gridlines
