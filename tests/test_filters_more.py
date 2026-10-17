"""Condition and colour filters, custom sort orders, advanced filter."""
import openpyxl

import data_more
import data_tools
import workbooks
from conftest import open_in


def table(app):
    rows = [["Name", "Qty", "Size"], ["a", "5", "Medium"], ["b", "50", "Low"], ["c", "15", "High"], ["d", "2", "Low"]]
    app.set_cells([(r, c, v) for r, row in enumerate(rows) for c, v in enumerate(row)])
    app.table.goto(0, 0)
    data_tools.toggle_filter(app)
    return app.table


def test_condition_filters(app):
    g = table(app)
    data_tools.set_criteria(app, 1, {"cond": [">", "10"]})
    assert g.sheet.filtered_rows == {1, 4}
    data_tools.set_criteria(app, 1, {"cond": ["top", "2"]})
    assert g.sheet.filtered_rows == {1, 4}
    data_tools.set_criteria(app, 1, {"cond": ["bottom", "1"]})
    assert g.sheet.filtered_rows == {1, 2, 3}
    data_tools.set_criteria(app, 1, None)
    data_tools.set_criteria(app, 2, {"cond": ["contains", "ow"]})
    assert g.sheet.filtered_rows == {1, 3}
    data_tools.set_criteria(app, 2, {"cond": ["begins", "h"]})
    assert g.sheet.filtered_rows == {1, 2, 4}


def test_colour_filter(app):
    g = table(app)
    g.select(2, 0, 3, 0)
    app.format_cells(lambda st: {**st, "bg": "#ffff00"})
    data_tools.set_criteria(app, 0, {"color": "#ffff00"})
    assert g.sheet.filtered_rows == {1, 4}
    data_tools.set_criteria(app, 0, {"color": ""})
    assert g.sheet.filtered_rows == {2, 3}


def test_popup_condition(app):
    g = table(app)
    p = data_tools.FilterPopup(app, 1, 0, 0)
    p.op.set("at least")
    p.want.insert(0, "15")
    p.apply_cond()
    assert g.sheet.filtered_rows == {1, 4}


def test_condition_filter_saved(app, tmp_path):
    path = workbooks.plain(tmp_path / "p.xlsx")
    open_in(app, path)
    table(app)
    data_tools.set_criteria(app, 1, {"cond": [">", "10"]})
    data_tools.set_criteria(app, 2, {"cond": ["contains", "ig"]})
    assert app.save()
    af = openpyxl.load_workbook(path).active.auto_filter
    ops = sorted((fc.customFilters.customFilter[0].operator, fc.customFilters.customFilter[0].val) for fc in af.filterColumn)
    assert ops == [("equal", "*ig*"), ("greaterThan", "10")]
    g = open_in(app, path)
    assert g.sheet.filter["criteria"] == {1: {"cond": [">", "10"]}, 2: {"cond": ["contains", "ig"]}}
    assert g.sheet.filtered_rows == {1, 2, 4}


def test_custom_sort_order(app):
    g = table(app)
    data_tools.set_criteria(app, 0, None)
    orders = data_tools.custom_orders()
    from simplecore import settings
    settings.set("fill_lists", [["Low", "Medium", "High"]])
    orders = data_tools.custom_orders()
    key = next(k for k in orders if k.startswith("Low"))
    data_tools.sort_range(app, (0, 0, 4, 2), [(2, True, orders[key])], header=True)
    assert [g.text(r, 2) for r in range(1, 5)] == ["Low", "Low", "Medium", "High"]


def test_criterion_and_advanced_filter(app):
    ok = data_more.criterion_ok
    assert ok("50", ">10") and not ok("5", ">10") and ok("East", "ea") and ok("Eastern", "Ea*")
    assert ok("West", "<>East") and ok("abc", "=a?c") and not ok("abd", "=abc")
    g = table(app)
    data_tools.set_criteria(app, 0, None)
    app.set_cells([(0, 5, "Qty"), (0, 6, "Size"), (1, 5, ">10"), (1, 6, "Low"), (2, 6, "Medium")])
    rows = data_more.advanced_filter(app, (0, 0, 4, 2), (0, 5, 2, 6))
    assert rows == [1, 2]                                   # (Qty>10 and Low) or Medium
    assert {3, 4} <= g.sheet.hidden_rows
    app.undo()
    rows = data_more.advanced_filter(app, (0, 0, 4, 2), (0, 5, 2, 6), copy_to=(10, 0))
    assert [g.text(r, 0) for r in (10, 11, 12)] == ["Name", "a", "b"]


def test_advanced_filter_menu(app, monkeypatch):
    g = table(app)
    data_tools.set_criteria(app, 0, None)
    app.set_cells([(0, 5, "Size"), (1, 5, "Low")])
    g.goto(1, 1)
    answers = iter(["F1:F2", "H1"])
    monkeypatch.setattr(data_more.simpledialog, "askstring", lambda *a, **k: next(answers))
    assert app.data_tool("advanced") == [2, 4]
    assert [g.text(r, 7) for r in range(3)] == ["Name", "b", "d"]
