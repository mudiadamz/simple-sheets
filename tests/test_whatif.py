"""Data → What-if: scenarios (saved in .xlsx), data tables, solver."""
import openpyxl
import pytest

import whatif
import workbooks
from conftest import open_in


def model(app):
    # B1 price, B2 qty, B3 = revenue
    app.set_cells([(0, 0, "Price"), (0, 1, "10"), (1, 0, "Qty"), (1, 1, "5"), (2, 0, "Revenue"), (2, 1, "=B1*B2")])
    return app.table


def test_cells_of():
    assert whatif.cells_of("B1, A2:B3") == [(0, 1), (1, 0), (1, 1), (2, 0), (2, 1)]
    assert whatif.cells_of("$C$4") == [(3, 2)]
    assert whatif.cells_of("nope!") is None


def test_scenarios_show_summary_and_undo(app):
    g = model(app)
    whatif.add_scenario(app, "Low", [(0, 1), (1, 1)], ["8", "4"])
    whatif.add_scenario(app, "High", [(0, 1), (1, 1)], ["12", "9"])
    assert [s["name"] for s in whatif.scenarios(app)] == ["Low", "High"]
    whatif.show_scenario(app, "High")
    assert g.value(2, 1) == "108"
    app.undo()
    assert g.value(2, 1) == "50"
    src = app.sheet
    name = whatif.summary(app, [(2, 1)])
    rows = app.sheets[name].rows
    assert rows[0] == ["", "Current values", "Low", "High"]
    assert rows[-1] == ["B3", "50", "32", "108"]
    assert app.sheets[src].rows[0][1] == "10"                       # trying scenarios left the sheet alone
    app.show(src)
    whatif.delete_scenario(app, "Low")
    assert [s["name"] for s in whatif.scenarios(app)] == ["High"]
    app.undo()
    assert len(whatif.scenarios(app)) == 2


def test_scenarios_saved(app, tmp_path):
    path = workbooks.plain(tmp_path / "s.xlsx")
    open_in(app, path)
    model(app)
    whatif.add_scenario(app, "Best", [(0, 1), (1, 1)], ["20", "7"])
    assert app.save()
    sl = openpyxl.load_workbook(path).active.scenarios
    assert sl.scenario[0].name == "Best" and [(c.r, c.val) for c in sl.scenario[0].inputCells] == [("B1", "20"), ("B2", "7")]
    open_in(app, path)
    assert whatif.scenarios(app) == [{"name": "Best", "comment": "", "cells": {"B1": "20", "B2": "7"}}]
    whatif.delete_scenario(app, "Best")
    assert app.save()
    assert not openpyxl.load_workbook(path).active.scenarios


def test_scenario_dialog(app):
    model(app)
    d = whatif.scenario_manager(app)
    assert d.add("Now", "B1:B2") == "Now" and d.list.get(0) == "Now"
    assert "B1 = 10" in d.info["text"]
    name = d.summary("B3")
    assert app.sheet == name


def test_one_variable_data_table(app):
    g = model(app)
    # D1 = formula (revenue), C2:C4 = prices to try in B1
    app.set_cells([(0, 3, "=B3"), (1, 2, "1"), (2, 2, "2"), (3, 2, "3")])
    assert whatif.data_table(app, (0, 2, 3, 3), col_input=(0, 1)) == 3
    assert [g.text(r, 3) for r in (1, 2, 3)] == ["5", "10", "15"]
    assert g.value(2, 1) == "50"


def test_two_variable_data_table(app):
    g = model(app)
    app.set_cells([(5, 0, "=B3"), (5, 1, "1"), (5, 2, "2"), (6, 0, "10"), (7, 0, "20")])
    whatif.data_table(app, (5, 0, 7, 2), row_input=(1, 1), col_input=(0, 1))   # qty across, price down
    assert [[g.text(r, c) for c in (1, 2)] for r in (6, 7)] == [["10", "20"], ["20", "40"]]


def test_data_table_dialog(app, monkeypatch):
    g = model(app)
    app.set_cells([(0, 3, "=B3"), (1, 2, "1"), (2, 2, "2")])
    g.select(0, 2, 2, 3)
    answers = iter(["", "B1"])
    monkeypatch.setattr(whatif.simpledialog, "askstring", lambda *a, **k: next(answers))
    assert whatif.data_table_dialog(app) == 2 and g.text(2, 3) == "10"


def test_constraints():
    cons = whatif.parse_constraints(["B1>=0", "B2:B3 <= 10", "B4=int", "C1<=D1", ""])
    assert cons[0] == ((0, 1), ">=", 0) and cons[2] == ((2, 1), "<=", 10)
    assert cons[3] == ((3, 1), "=", "int") and cons[4] == ((0, 2), "<=", (0, 3))
    with pytest.raises(ValueError):
        whatif.parse_constraints(["B1 >> 2"])


def test_solver_max_with_constraints(app):
    g = model(app)
    # maximise revenue with price <= 15, qty <= 8 - price/5 style trade-off: B2 = 20 - B1 (via constraint)
    app.set_cells([(1, 1, "=20-B1")])
    x, obj, ok = whatif.solve(app, (2, 1), "max", [(0, 1)], whatif.parse_constraints(["B1<=15", "B1>=0"]))
    assert ok and abs(x[0] - 10) < 1e-3 and abs(obj - 100) < 1e-4
    assert g.text(0, 1) == "10"                                      # solve() itself leaves the sheet alone


def test_solver_value_and_int(app):
    model(app)
    x, obj, ok = whatif.solve(app, (2, 1), "value", [(1, 1)], whatif.parse_constraints(["B2=int"]), value=70)
    assert ok and x == [7] and obj == 70


def test_solver_dialog(app):
    g = model(app)
    d = whatif.solver(app)
    d.target.delete(0, "end")
    d.target.insert(0, "B3")
    d.goal.set("min")
    d.changing.insert(0, "B1")
    d.cons.insert("1.0", "B1>=3\nB1<=9")
    assert d.ok() is True
    assert g.text(0, 1) == "3" and g.value(2, 1) == "15"
    app.undo()
    assert g.text(0, 1) == "10"
