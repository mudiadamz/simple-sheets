"""Clicking cells while typing a formula inserts their references (Excel 'point' mode)."""
import workbooks
from conftest import key, open_in


def cell_xy(g, r, c):
    return g.col_x(c) + 5, g.row_y(r) + 5


def click(app, r, c, shift=False):
    g = app.table
    x, y = cell_xy(g, r, c)
    g.cv.event_generate("<Shift-Button-1>" if shift else "<Button-1>", x=x, y=y, when="now")
    g.cv.event_generate("<ButtonRelease-1>", x=x, y=y, when="now")
    app.update()


def start_formula(app, r, c, text="="):
    g = app.table
    g.goto(r, c)
    g.edit(text)
    app.update()
    return g.editor


def test_click_inserts_and_replaces_reference(app, tmp_path):
    g = open_in(app, workbooks.budget(tmp_path / "b.xlsx"))
    e = start_formula(app, 7, 0)
    click(app, 1, 1)
    assert e.get() == "=B2" and g.cur == (7, 0)          # selection stays on the edited cell
    assert [m[:4] for m in g.ref_marks] == [(1, 1, 1, 1)]
    click(app, 2, 2)
    assert e.get() == "=C3"                               # a new click replaces the pointed ref
    click(app, 3, 3, shift=True)
    assert e.get() == "=C3:D4"                            # shift-click makes a range
    assert [m[:4] for m in g.ref_marks] == [(2, 2, 3, 3)]


def test_operator_then_click_appends(app, tmp_path):
    g = open_in(app, workbooks.budget(tmp_path / "b.xlsx"))
    e = start_formula(app, 7, 0)
    click(app, 1, 1)
    e.insert("end", "*")
    e.icursor("end")
    click(app, 1, 2)
    assert e.get() == "=B2*C2"
    key(app, "<Return>", widget=e)
    assert g.text(7, 0) == "=B2*C2" and g.value(7, 0) == "25" and g.ref_marks == []


def test_function_args_and_drag_range(app, tmp_path):
    g = open_in(app, workbooks.budget(tmp_path / "b.xlsx"))
    e = start_formula(app, 7, 0, "=SUM(")
    x1, y1 = cell_xy(g, 1, 3)
    x2, y2 = cell_xy(g, 3, 3)
    g.cv.event_generate("<Button-1>", x=x1, y=y1, when="now")
    g.cv.event_generate("<B1-Motion>", x=x2, y=y2, when="now")
    g.cv.event_generate("<ButtonRelease-1>", x=x2, y=y2, when="now")
    app.update()
    assert e.get() == "=SUM(D2:D4"


def test_header_click_inserts_whole_column_or_row(app, tmp_path):
    g = open_in(app, workbooks.budget(tmp_path / "b.xlsx"))
    e = start_formula(app, 7, 0, "=SUM(")
    g.cv.event_generate("<Button-1>", x=g.col_x(1) + 5, y=3, when="now")
    app.update()
    assert e.get() == "=SUM(B:B"
    g.cv.event_generate("<Shift-Button-1>", x=g.col_x(2) + 5, y=3, when="now")
    app.update()
    assert e.get() == "=SUM(B:C"
    g.cv.event_generate("<Button-1>", x=5, y=g.row_y(2) + 5, when="now")
    app.update()
    assert e.get() == "=SUM(3:3"


def test_click_after_complete_reference_finishes_edit(app, tmp_path):
    g = open_in(app, workbooks.budget(tmp_path / "b.xlsx"))
    start_formula(app, 7, 0, "=B2")   # caret after "B2": not a place for a new reference
    click(app, 2, 2)
    assert g.editor is None and g.text(7, 0) == "=B2" and g.cur == (2, 2)


def test_plain_text_click_finishes_edit(app, tmp_path):
    g = open_in(app, workbooks.budget(tmp_path / "b.xlsx"))
    start_formula(app, 7, 0, "hello")
    click(app, 2, 2)
    assert g.editor is None and g.text(7, 0) == "hello"


def test_point_from_formula_bar(app, tmp_path):
    g = open_in(app, workbooks.budget(tmp_path / "b.xlsx"))
    g.goto(7, 0)
    v = app.value
    v.focus_force()
    for _ in range(50):
        app.update()
        if app.focus_get() is v:
            break
    v.delete(0, "end")
    v.insert(0, "=")
    v.icursor("end")
    click(app, 1, 1)
    click(app, 3, 1, shift=True)
    assert v.get() == "=B2:B4" and g.cur == (7, 0)
    app.value_commit()
    assert g.text(7, 0) == "=B2:B4"


class Wheel:
    def __init__(self, delta):
        self.delta, self.x, self.y = delta, 5, 5


def test_scroll_while_typing_formula_keeps_editing(app, tmp_path):
    """Typing =, scrolling down with the wheel to a far cell, clicking it: the edit stays open."""
    g = open_in(app, workbooks.budget(tmp_path / "b.xlsx"))
    app.set_cells([(80, 1, "42")])
    e = start_formula(app, 2, 0, "=SUM(")
    for _ in range(25):                                   # 75 rows down: the edited cell is off screen
        g.on_wheel(Wheel(-120))
    app.update()
    assert g.editor is e and e.winfo_exists() and g.top >= 70
    assert int(e.place_info()["y"]) < 0                   # waiting out of sight, still the editor
    click(app, 80, 1)
    assert e.get() == "=SUM(B81"
    e.focus_force()
    app.update()
    e.event_generate("<KeyPress>", keysym="plus", when="now")   # typing still reaches the hidden editor
    assert app.focus_get() is e and e.get() == "=SUM(B81+"
    e.delete(len(e.get()) - 1, "end")
    e.insert("end", ")")
    for _ in range(25):                                   # back up: the editor sits on its cell again
        g.on_wheel(Wheel(120))
    app.update()
    assert int(e.place_info()["y"]) == g.row_y(2) and int(e.place_info()["x"]) == g.col_x(0)
    key(app, "<Return>", e)
    assert g.text(2, 0) == "=SUM(B81)" and g.value(2, 0) == "42"


def test_scrollbar_while_typing_formula(app, tmp_path):
    g = open_in(app, workbooks.budget(tmp_path / "b.xlsx"))
    e = start_formula(app, 1, 1)
    g.yview("scroll", 40, "units")
    g.ys.focus_set()                                      # a click on the scrollbar takes the focus
    app.update()
    assert g.editor is e and app.focus_get() is e
    g.commit(cancel=True)


def test_wheel_still_finishes_plain_text(app, tmp_path):
    g = open_in(app, workbooks.budget(tmp_path / "b.xlsx"))
    e = start_formula(app, 1, 1, "hello")
    g.on_wheel(Wheel(-120))
    assert g.editor is None and not e.winfo_exists() and g.text(1, 1) == "hello"
