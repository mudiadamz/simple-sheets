"""Row height: drag the row number's bottom edge; double-click for automatic."""
from types import SimpleNamespace

import openpyxl

import sheet_io as io
import workbooks
from conftest import open_in


def drag_row(app, r, dy):
    g = app.table
    app.update()
    y = g.row_y(r) + g.rh(r)
    x = g.hdr_w() // 2
    g.cv.event_generate("<ButtonPress-1>", x=x, y=y, when="now")
    g.cv.event_generate("<B1-Motion>", x=x, y=y + dy, when="now")
    g.cv.event_generate("<ButtonRelease-1>", x=x, y=y + dy, when="now")
    app.update()


def test_drag_taller_and_shorter_with_undo(app):
    g = app.table
    h0 = g.rh(2)
    drag_row(app, 2, g.px(30))
    assert abs(g.rh(2) - (h0 + g.px(30))) <= 1 and 2 in g.sheet.heights and app.dirty()   # stored in points
    drag_row(app, 2, -g.px(40))
    assert g.rh(2) < h0                                    # rows can be made smaller than default
    app.undo()
    assert abs(g.rh(2) - (h0 + g.px(30))) <= 1
    app.undo()
    assert g.rh(2) == h0 and not app.dirty()


def test_double_click_back_to_automatic(app):
    g = app.table
    app.set_height(1, 50)
    app.update()
    y = g.row_y(1) + g.rh(1)
    g.on_double(SimpleNamespace(x=g.hdr_w() // 2, y=y))
    assert 1 not in g.sheet.heights and g.rh(1) == g.rowh


def test_saved_and_loaded(app, tmp_path):
    path = workbooks.plain(tmp_path / "p.xlsx")
    open_in(app, path)
    app.set_height(0, 33.5)
    assert app.save(), app.dialogs
    assert openpyxl.load_workbook(path)["Data"].row_dimensions[1].height == 33.5
    assert io.load(str(path))["Data"].heights[0] == 33.5
    app.set_height(0, None)
    assert app.save(), app.dialogs
    assert not openpyxl.load_workbook(path)["Data"].row_dimensions[1].height


def test_cursor_over_row_edge(app):
    g = app.table
    app.update()
    y = g.row_y(3) + g.rh(3)
    g.on_motion(SimpleNamespace(x=g.hdr_w() // 2, y=y))
    assert g.cv["cursor"] == "sb_v_double_arrow"
