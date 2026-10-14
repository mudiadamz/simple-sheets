"""Number/date formats and reading styles from files."""
import pytest

import sheet_io as io
import workbooks
from formatting import render


@pytest.mark.parametrize("text, fmt, want", [
    ("1234.5", "0.00", "1234.50"),
    ("1234.5", "#,##0", "1,235"),
    ("1234567.891", "#,##0.00", "1,234,567.89"),
    ("0.256", "0%", "26%"),
    ("0.256", "0.0%", "25.6%"),
    ("-1234.5", '"$"#,##0.00', "-$1,234.50"),
    ("-1234.5", "#,##0.00;[Red](#,##0.00)", "(1,234.50)"),
    ("1234.5", "[$€-407]#,##0.00", "€1,234.50"),
    ("0", '#,##0;-#,##0;"-"', "-"),
    ("12345", "0.00E+00", "1.23E+04"),
    ("0.5", "#.00", ".50"),
    ("7", "000", "007"),
    ("2.5", "0.##", "2.5"),
    ("3", "0.##", "3."),
    ("1500000", '#,##0,,"M"', "2M"),
    ("2024-03-05", "dd/mm/yyyy", "05/03/2024"),
    ("2024-03-05", "mmm d, yyyy", "Mar 5, 2024"),
    ("2024-03-05 14:07:09", "h:mm AM/PM", "2:07 PM"),
    ("2024-03-05 14:07:09", "yyyy-mm-dd hh:mm:ss", "2024-03-05 14:07:09"),
    ("45356", "d-mmm-yy", "5-Mar-24"),
    ("hello", "0.00", "hello"),
    ("1.5", "General", "1.5"),
    ("0.333333", "_(* #,##0.00_)", " 0.33 "),
])
def test_number_formats(text, fmt, want):
    assert render(text, fmt) == want


def test_load_styles(tmp_path):
    sh = io.load(str(workbooks.styled(tmp_path / "styled.xlsx")))["Report"]
    assert sh.rows[2] == ["Apple", "1200", "0.5", "2024-03-05"]
    assert sh.merges == [(0, 0, 0, 3)]
    assert sh.widths == {0: 25.0}
    assert sh.styles[(0, 0)] == {"bold": True, "fg": "#ffffff", "bg": "#376092", "align": "center"}
    assert sh.styles[(2, 0)] == {"italic": True, "fg": "#ff0000"}
    shown = {k: render(sh.rows[k[0]][k[1]], sh.styles[k]["numfmt"]) for k in [(2, 1), (2, 2), (2, 3)]}
    assert shown == {(2, 1): "1,200", (2, 2): "$0.50", (2, 3): "Mar 5, 2024"}


def test_load_csv_and_ods(tmp_path):
    p = tmp_path / "semi.csv"
    p.write_text("a;b\n1;2\n", encoding="utf-8")
    assert io.load(str(p))["semi.csv"].rows == [["a", "b"], ["1", "2"]]

    import pandas as pd   # only used to produce a real .ods file
    pytest.importorskip("odf")
    ods = tmp_path / "t.ods"
    pd.DataFrame({"a": [1, 2], "b": ["x", None]}).to_excel(ods, engine="odf", index=False)
    assert io.load(str(ods))["Sheet1"].rows == [["a", "b"], ["1", "x"], ["2"]]


def test_csv_save_keeps_delimiter(app, tmp_path):
    from conftest import open_in
    p = tmp_path / "e.csv"
    p.write_text("a;b\n1;2\n", encoding="utf-8")
    g = open_in(app, p)
    g.goto(1, 1)
    app.set_cells([(1, 1, "5")])
    assert app.save()
    assert p.read_text(encoding="utf-8").splitlines() == ["a;b", "1;5"]
