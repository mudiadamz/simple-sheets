"""Tools → Compare with another workbook."""
import openpyxl

import compare
from conftest import open_in


def two_books(tmp_path):
    a, b = tmp_path / "a.xlsx", tmp_path / "b.xlsx"
    for path, rows, extra in ((a, [["x", 1, 2], ["y", 3, "=B2*2"]], "OnlyA"),
                              (b, [["x", 1.0, 5], ["y", 3, "=B2*3"], ["new"]], "OnlyB")):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Data"
        for row in rows:
            ws.append(row)
        wb.create_sheet(extra)
        wb.save(path)
    return a, b


def test_diff_sheet():
    assert compare.diff_sheet([["a", "1"]], [["a", "1.0", "x"]]) == [(0, 2, "", "x")]   # 1 == 1.0
    assert compare.diff_sheet([["a"], ["b"]], [["a"]]) == [(1, 0, "b", "")]


def test_compare_books(app, tmp_path):
    a, b = two_books(tmp_path)
    open_in(app, a)
    res = compare.compare_file(app, str(b))
    assert res["only_mine"] == ["OnlyA"] and res["only_other"] == ["OnlyB"]
    assert res["cells"] == [("Data", 0, 2, "2", "5"), ("Data", 1, 2, "=B2*2", "=B2*3"), ("Data", 2, 0, "", "new")]


def test_window_go_and_report(app, tmp_path):
    a, b = two_books(tmp_path)
    open_in(app, a)
    w = compare.open_compare(app, str(b))
    assert len(w.tree.get_children()) == 3 and "3 cells differ" in w.summary["text"]
    app.show("OnlyA")
    assert w.go(1) == ("Data", 1, 2) and app.sheet == "Data" and app.table.cur == (1, 2)
    name = w.report()
    rows = app.sheets[name].rows
    assert rows[1] == ["Data", "C1", "2", "5"] and rows[-1] == ["OnlyB", "(sheet only in the other file)", "", ""]
    app.undo()
    assert name not in app.sheets


def test_bad_file(app, tmp_path):
    bad = tmp_path / "bad.xlsx"
    bad.write_text("nope")
    assert compare.open_compare(app, str(bad)) is None and app.dialogs[-1][0] == "showerror"
