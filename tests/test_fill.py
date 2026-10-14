"""Fill handle: series rules and the drag / double-click / Ctrl+D / Ctrl+R interactions."""
import pytest

import workbooks
from conftest import key, open_in
from fill import extend


def texts(values, n, copy=False):
    return [t if t is not None else values[i] for i, t in extend(values, n, copy)]


@pytest.mark.parametrize("values, n, want", [
    (["1"], 3, ["2", "3", "4"]),                               # one number counts up
    (["2", "4"], 3, ["6", "8", "10"]),                          # step from the pattern
    (["10", "7"], 2, ["4", "1"]),
    (["0.5", "1"], 2, ["1.5", "2"]),
    (["1.5"], 2, ["2.5", "3.5"]),
    (["2024-01-30"], 3, ["2024-01-31", "2024-02-01", "2024-02-02"]),
    (["2024-01-01", "2024-01-08"], 2, ["2024-01-15", "2024-01-22"]),
    (["Monday"], 3, ["Tuesday", "Wednesday", "Thursday"]),
    (["Sat"], 2, ["Sun", "Mon"]),
    (["JAN"], 2, ["FEB", "MAR"]),
    (["November"], 3, ["December", "January", "February"]),
    (["Item 1"], 2, ["Item 2", "Item 3"]),
    (["A-007"], 2, ["A-008", "A-009"]),
    (["Week 1", "Week 3"], 2, ["Week 5", "Week 7"]),
    (["apple", "pear"], 3, ["apple", "pear", "apple"]),         # plain text repeats
    (["1", "x"], 2, ["1", "x"]),                                # mixed: repeat
    (["", ""], 2, ["", ""]),
])
def test_series(values, n, want):
    assert texts(values, n) == want


def test_copy_flag_and_formulas_repeat():
    assert texts(["1"], 3, copy=True) == ["1", "1", "1"]
    assert extend(["=A1"], 2) == [(0, None), (0, None)]          # caller shifts formulas


def drag_handle(app, to_r, to_c, ctrl=False):
    g = app.table
    app.update()
    hx, hy = g.handle
    g.cv.event_generate("<ButtonPress-1>", x=hx, y=hy, when="now")
    x, y = g.col_x(to_c) + 5, g.row_y(to_r) + 5
    g.cv.event_generate("<B1-Motion>", x=x, y=y, when="now")
    app.update()
    preview = g.fill_preview
    g.cv.event_generate("<ButtonRelease-1>", x=x, y=y, state=0x4 if ctrl else 0, when="now")
    app.update()
    return preview


def test_drag_handle_counts_numbers(app):
    g = app.table
    app.set_cells([(0, 0, "1")])
    g.goto(0, 0)
    preview = drag_handle(app, 4, 0)
    assert preview == (1, 0, 4, 0)
    assert [g.text(r, 0) for r in range(5)] == ["1", "2", "3", "4", "5"]
    assert g.sel() == (0, 0, 4, 0) and g.fill_preview is None
    app.undo()                                                   # one undo step
    assert [g.text(r, 0) for r in range(5)] == ["1", "", "", "", ""]


def test_ctrl_drag_copies(app):
    g = app.table
    app.set_cells([(0, 0, "1")])
    g.goto(0, 0)
    drag_handle(app, 2, 0, ctrl=True)
    assert [g.text(r, 0) for r in range(3)] == ["1", "1", "1"]


def test_drag_right_and_up(app):
    g = app.table
    app.set_cells([(5, 1, "Mon"), (5, 2, "Tue")])
    g.select(5, 1, 5, 2)
    drag_handle(app, 5, 4)
    assert [g.text(5, c) for c in range(1, 5)] == ["Mon", "Tue", "Wed", "Thu"]
    app.set_cells([(5, 6, "10")])
    g.goto(5, 6)
    app.update()
    hx, hy = g.handle
    g.cv.event_generate("<ButtonPress-1>", x=hx, y=hy, when="now")
    g.cv.event_generate("<B1-Motion>", x=hx - 2, y=g.row_y(3) + 5, when="now")
    g.cv.event_generate("<ButtonRelease-1>", x=hx - 2, y=g.row_y(3) + 5, when="now")
    app.update()
    assert [g.text(r, 6) for r in (3, 4, 5)] == ["12", "11", "10"]   # upward continues backwards


def test_fill_formulas_and_formats(app, tmp_path):
    g = open_in(app, workbooks.budget(tmp_path / "b.xlsx"))
    app.set_cells([(4, 0, "Tape"), (4, 1, "4"), (4, 2, "1.25")])
    g.goto(3, 3)                                                # D4 = B4*C4, currency format
    drag_handle(app, 4, 3)
    assert g.text(4, 3) == "=B5*C5" and g.value(4, 3) == "5"
    assert g.sheet.styles.get((4, 3), {}).get("numfmt") == '"$"#,##0.00'


def test_double_click_fills_to_neighbour(app):
    g = app.table
    app.set_cells([(r, 0, f"name{r}") for r in range(6)] + [(0, 1, "1")])
    g.goto(0, 1)
    app.update()
    assert g.cv.bind("<Double-1>")
    from types import SimpleNamespace
    hx, hy = g.handle
    g.on_double(SimpleNamespace(x=hx, y=hy))
    assert [g.text(r, 1) for r in range(7)] == ["1", "2", "3", "4", "5", "6", ""]


def test_ctrl_d_and_ctrl_r(app):
    g = app.table
    app.set_cells([(0, 0, "x"), (0, 1, "=A1&\"!\"")])
    g.select(0, 0, 3, 1)
    key(app, "<Control-d>")
    assert [g.text(r, 0) for r in range(4)] == ["x"] * 4
    assert g.text(3, 1) == '=A4&"!"' and g.value(3, 1) == "x!"
    g.goto(0, 2)
    key(app, "<Control-r>")                                       # single cell: copy from the left
    assert g.text(0, 2) == '=B1&"!"'
