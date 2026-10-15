"""The showcase workbook (tools/make_showcase.py) opens and shows its formulas, rules and charts; chart axes."""
import os
import sys

import charts
from conftest import open_in
from formatting import render

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))
import make_showcase  # noqa: E402


def test_showcase_opens(app, tmp_path):
    path = make_showcase.build(str(tmp_path / "showcase.xlsx"))
    g = open_in(app, path)
    assert list(app.sheets) == ["Dashboard", "Orders", "Roadmap"]
    shown = lambda r, c: render(g.value(r, c), (g.sheet.styles.get((r, c)) or {}).get("numfmt"))  # noqa: E731
    assert shown(5, 1) == "$519,900"                           # KPI card: formula + number format
    assert shown(24, 5) == "37.7%" and len(g.sheet.charts) == 3 and g.sheet.cond
    app.show("Orders")
    assert app.table.sheet.extra.get("tables") and app.table.value(1, 7) == "130.5"    # =Qty*Price
    app.show("Roadmap")
    assert app.table.value(3, 4) == "0.9"


def test_nice_axis():
    assert charts.nice_scale(0, 54800) == (0, 60000, 20000)
    assert charts.nice_scale(-3, 7) == (-5, 7.5, 2.5)
    assert charts.nice_scale(5, 5)[1] > 5
    assert charts.axis_label(54800) == "54.8k" and charts.axis_label(1_500_000) == "1.5M"
    assert charts.axis_label(0.25) == "0.25" and charts.axis_label(-20000) == "-20k"
