"""Sort and filter."""
import openpyxl

import data_tools
import sheet_io as io
from conftest import open_in

DATA = [["Fruit", "Qty", "Total"],
        ["pear", "5", "=B2*2"],
        ["Apple", "12", "=B3*2"],
        ["fig", "", "=B4*2"],
        ["banana", "3", "=B5*2"]]


def load(app):
    app.set_cells([(r, c, v) for r, row in enumerate(DATA) for c, v in enumerate(row)])
    app.table.select(0, 0, 0, 2)
    app.format_cells(lambda st: {**st, "bold": True})
    return app.table


def col(g, c, rows=range(1, 5)):
    return [g.text(r, c) for r in rows]


def test_quick_sort_keeps_rows_and_moves_formulas(app):
    g = load(app)
    g.goto(2, 0)
    assert data_tools.quick_sort(app, True)
    assert col(g, 0) == ["Apple", "banana", "fig", "pear"]           # header kept, case-insensitive
    assert col(g, 2) == ["=B2*2", "=B3*2", "=B4*2", "=B5*2"]          # formulas follow their row
    assert [g.value(r, 2) for r in range(1, 5)] == ["24", "6", "0", "10"]
    assert g.sheet.styles.get((0, 0), {}).get("bold") and not g.sheet.styles.get((1, 0))
    app.undo()
    assert col(g, 0) == ["pear", "Apple", "fig", "banana"]


def test_numbers_descending_blanks_last(app):
    g = load(app)
    g.goto(1, 1)
    data_tools.quick_sort(app, False)
    assert col(g, 1) == ["12", "5", "3", ""]
    data_tools.quick_sort(app, True)
    assert col(g, 1) == ["3", "5", "12", ""]                          # blanks last both ways


def test_sort_dialog_two_keys(app):
    g = app.table
    app.set_cells([(r, c, v) for r, row in enumerate(
        [["Team", "Score"], ["b", "1"], ["a", "2"], ["b", "3"], ["a", "1"]]) for c, v in enumerate(row)])
    g.goto(1, 0)
    d = data_tools.SortDialog(app)
    assert d.header.get()                                            # header detected
    d.cols[0].set(d.labels[0])
    d.cols[1].set(d.labels[1])
    d.orders[1].set("Z → A")
    d.ok()
    assert [g.text(r, 0) + g.text(r, 1) for r in range(1, 5)] == ["a2", "a1", "b3", "b1"]


def test_filter_create_hide_and_remove(app):
    g = load(app)
    g.goto(0, 0)
    data_tools.toggle_filter(app)
    assert g.sheet.filter["range"] == (0, 0, 4, 2)
    assert set(g.filter_buttons) == {0, 1, 2}
    assert data_tools.values_in(app, app.sheet, 1) == ["3", "5", "12", "(Blanks)"]
    data_tools.set_criteria(app, 1, ["5", "12"])
    assert g.sheet.filtered_rows == {3, 4} and g.rh(3) == 0 and g.rh(1) > 0
    data_tools.set_criteria(app, 0, ["pear"])
    assert g.sheet.filtered_rows == {2, 3, 4}
    app.undo()
    assert g.sheet.filtered_rows == {3, 4}
    data_tools.toggle_filter(app)                                     # remove
    assert g.sheet.filter is None and not g.sheet.filtered_rows


def test_filter_popup(app):
    g = load(app)
    g.goto(0, 0)
    data_tools.toggle_filter(app)
    app.update()
    x1, y1, x2, y2 = g.filter_buttons[0]
    g.cv.event_generate("<Button-1>", x=(x1 + x2) // 2, y=(y1 + y2) // 2, when="now")
    app.update()
    p = app.filter_popup
    assert list(p.vars) == ["Apple", "banana", "fig", "pear"]
    p.vars["fig"].set(False)
    p.ok()
    assert g.sheet.filtered_rows == {3}
    assert g.sheet.filter["criteria"][0] == ["Apple", "banana", "pear"]


def test_filter_saved_and_loaded(app, tmp_path, monkeypatch):
    import simple_sheets
    dst = tmp_path / "f.xlsx"
    monkeypatch.setattr(simple_sheets.filedialog, "asksaveasfilename", lambda **k: str(dst))
    g = load(app)
    g.goto(0, 0)
    data_tools.toggle_filter(app)
    data_tools.set_criteria(app, 1, ["5", "12", "(Blanks)"])
    assert app.save(), app.dialogs
    ws = openpyxl.load_workbook(dst).active
    assert ws.auto_filter.ref == "A1:C5"
    assert {r for r, d in ws.row_dimensions.items() if d.hidden} == {5}      # banana (3) hidden
    s = io.load(str(dst))["Sheet1"]
    assert s.filter == {"range": (0, 0, 4, 2), "criteria": {1: ["5", "12", "(Blanks)"]}}
    g = open_in(app, dst)
    assert g.sheet.filtered_rows == {4} and not g.sheet.hidden_rows        # filtered, not hand-hidden


def test_filter_follows_inserted_rows(app):
    g = load(app)
    g.goto(0, 0)
    data_tools.toggle_filter(app)
    data_tools.set_criteria(app, 0, ["pear"])
    app.restructure("row", 0, 2)
    assert g.sheet.filter["range"] == (2, 0, 6, 2) and g.sheet.filtered_rows == {4, 5, 6}


def test_menu_items(app):
    from test_menu import menu_states
    load(app)
    states = menu_states(app)
    assert {"Sort A → Z", "Sort Z → A", "Sort range…", "Create filter"} <= set(states)
