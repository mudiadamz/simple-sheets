"""A #CIRC! cell: the status bar says which cells go round in a circle (here: a pivot count using its own cell)."""
import workbooks
from conftest import open_in


def test_status_bar_names_the_loop(app, tmp_path):
    open_in(app, workbooks.circular_pivot(tmp_path / "pivot.xlsx"))
    app.show("Pivot")
    app.update()
    g = app.table
    assert app.engine.display(g.sheet, 2, 1) == "#CIRC!"
    g.goto(2, 1)                                                   # B3
    app.on_select()
    assert app.status["text"].startswith("Circular reference: B3 uses itself")
    g.goto(2, 3)                                                   # D3 = SUM(B3:C3)
    app.on_select()
    assert "(this cell uses B3)" in app.status["text"]
    g.goto(1, 1)                                                   # a header: no message
    app.on_select()
    assert "Circular" not in app.status["text"]
