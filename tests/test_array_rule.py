"""Excel's array rule: one-value functions given a range work value by value (TEXT(A2:A500,"yyyy-mm") inside
SUMPRODUCT, IF(range>5, …) inside SUM, --(range>x)…), as in pivotable_sales.xlsx's monthly summary."""
import openpyxl
import pytest

from conftest import open_in
from formula import Engine
from sheet_io import Sheet


@pytest.fixture
def ev():
    data = Sheet([["2024-01-05", "10", "apple"], ["2024-02-07", "20", "pear"], ["2024-01-20", "30", "fig"],
                  ["", "x", ""]])
    eng = Engine()
    eng.set_book({"S": data})
    eng.set_names({"FACTO": "=LAMBDA(n, IF(n<=1, 1, n*FACTO(n-1)))"})

    def run(formula):
        data.rows.append([formula])
        eng.clear()
        try:
            return eng.display(data, len(data.rows) - 1, 0)
        finally:
            data.rows.pop()
    return run


@pytest.mark.parametrize("formula, want", [
    ('=SUMPRODUCT((TEXT(A1:A3,"yyyy-mm")="2024-01")*B1:B3)', "40"),
    ("=SUMPRODUCT((MONTH(A1:A3)=1)*B1:B3)", "40"), ("=SUMPRODUCT((YEAR(A1:A3)=2024)*1)", "3"),
    ("=SUMPRODUCT(LEN(C1:C3))", "12"), ('=SUMPRODUCT((LEFT(C1:C3,1)="a")*B1:B3)', "10"),
    ("=SUMPRODUCT(ROUND(B1:B3/7,0))", "8"), ("=SUMPRODUCT(ABS(B1:B3-25))", "25"),
    ('=SUMPRODUCT(--ISNUMBER(SEARCH("p",C1:C3)))', "2"), ("=SUMPRODUCT(DAY(A1:A3))", "32"),
    ('=SUM(IF(TEXT(A1:A3,"mm")="01",B1:B3,0))', "40"), ("=SUMPRODUCT(--(A1:A3>DATE(2024,1,10)))", "2"),
    ("=SUMPRODUCT(-B1:B3)", "-60"), ("=SUMPRODUCT(N(B1:B3))", "60"), ("=SUM(IFERROR(B1:B4*1,0))", "60"),
    ('=SUMPRODUCT(--ISBLANK(A1:A4))', "1"), ("=SUMPRODUCT(EOMONTH(A1:A3,0)-A1:A3)", "59"),
    ("=CONCAT(C1:C3)", "applepearfig"),                       # whole ranges stay whole where they should
    ('=TEXTJOIN("-",TRUE,C1:C3)', "apple-pear-fig"), ("=SUM(B1:B3)", "60"), ("=COUNTIF(B1:B3,\">15\")", "2"),
    ('=TEXT(A1,"yyyy-mm")', "2024-01"), ("=LEN(C2)", "4"),     # one cell: as before
    ("=FACTO(5)", "120"),                                      # IF stays lazy: recursion still ends
    ('=IF(B4="x","skip",1/0)', "skip"), ("=IFERROR(1/0,7)", "7"),
    ("=N(TRUE)", "1"), ('=N("abc")', "0"), ('=T("abc")', "abc"), ("=T(5)", ""),
])
def test_value_by_value(ev, formula, want):
    assert ev(formula) == want


def test_spills_as_a_block(app):
    app.set_cells([(0, 0, "a"), (1, 0, "bb"), (2, 0, "ccc"), (0, 2, "=LEN(A1:A3)")])
    g = app.table
    assert [g.value(r, 2) for r in range(3)] == ["1", "2", "3"]   # dynamic array: one value per row


def test_monthly_summary_like_pivotable_sales(app, tmp_path):
    wb = openpyxl.Workbook()
    data = wb.active
    data.title = "Sheet1"
    data.append(["Date", "Salesperson", "Product", "Amount"])
    import datetime as dt
    rows = [(dt.datetime(2024, 1, 3), 100.5), (dt.datetime(2024, 1, 28), 50), (dt.datetime(2024, 2, 2), 75)]
    for d, amount in rows:
        data.append([d, "Bob", "Widget", amount])
    piv = wb.create_sheet("Pivot")
    piv["A1"], piv["B1"], piv["C1"] = "Month", "Total Sales", "Transactions"
    for i, m in enumerate(("2024-01", "2024-02"), start=2):
        piv[f"A{i}"] = m
        piv[f"B{i}"] = f'=SUMPRODUCT((TEXT(Sheet1!$A$2:$A$4,"yyyy-mm")=$A{i})*Sheet1!$D$2:$D$4)'
        piv[f"C{i}"] = f'=SUMPRODUCT((TEXT(Sheet1!$A$2:$A$4,"yyyy-mm")=$A{i})*1)'
    path = tmp_path / "sales.xlsx"
    wb.save(path)
    open_in(app, path)
    app.show("Pivot")
    g = app.table
    assert [g.value(1, 1), g.value(1, 2), g.value(2, 1), g.value(2, 2)] == ["150.5", "2", "75", "1"]
