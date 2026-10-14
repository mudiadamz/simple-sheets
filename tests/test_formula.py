"""Formula engine: operators, functions, errors, references."""
import time

import pytest

from formula import Engine, adjust, shift
from sheet_io import Sheet


@pytest.fixture
def ev():
    data = Sheet([
        ["Item", "Qty", "Price", "Date", "Flag"],
        ["Apple", "3", "1.5", "2024-03-05", "TRUE"],
        ["Pear", "5", "2", "2024-03-10", "FALSE"],
        ["Plum", "", "0.5", "", "x"],
        ["Fig", "10", "abc", "", ""],
    ])
    other = Sheet([["k", "v"], ["a", "100"], ["b", "200"]])
    eng = Engine()
    eng.set_book({"Data": data, "My Sheet": other})

    def run(formula):
        data.rows.append([formula])
        eng.clear()
        try:
            return eng.display(data, len(data.rows) - 1, 0)
        finally:
            data.rows.pop()
    return run


@pytest.mark.parametrize("formula, want", [
    ("=1+2*3", "7"), ("=(1+2)*3", "9"), ("=2^3^2", "64"), ("=-2^2", "4"), ("=10/4", "2.5"), ("=50%", "0.5"),
    ("=1/0", "#DIV/0!"), ('="a"&"b"&1', "ab1"), ("=0.1+0.2", "0.3"), ("=1=1", "TRUE"), ('="A"="a"', "TRUE"),
    ("=3>2", "TRUE"), ('=1<"a"', "TRUE"),
    ("=SUM(B2:B5)", "18"), ("=SUM(B:B)", "18"), ("=AVERAGE(B2:B5)", "6"), ("=MIN(C2:C5)", "0.5"),
    ("=MAX(B2:B5,20)", "20"), ("=COUNT(B2:C5)", "6"), ("=COUNTA(A2:A5)", "4"), ("=COUNTBLANK(B2:B5)", "1"),
    ("=B2*C2", "4.5"), ("=C5*2", "#VALUE!"), ("=SUM(C2:C5)", "4"), ("=PRODUCT(B2:B3)", "15"),
    ("=ROUND(2.345,2)", "2.35"), ("=ROUND(2.5,0)", "3"), ("=ROUND(-2.5,0)", "-3"), ("=ROUND(1234,-2)", "1200"),
    ("=ROUNDUP(1.21,1)", "1.3"), ("=ROUNDDOWN(1.29,1)", "1.2"), ("=INT(-1.5)", "-2"), ("=MOD(-3,2)", "1"),
    ("=SQRT(-1)", "#NUM!"), ("=ABS(-4)", "4"), ("=POWER(2,10)", "1024"),
    ('=IF(B2>2,"big","small")', "big"), ("=IF(B4,1,2)", "2"), ('=IFERROR(1/0,"oops")', "oops"),
    ("=IFERROR(5,0)", "5"), ("=AND(TRUE,1)", "TRUE"), ("=OR(FALSE,0)", "FALSE"), ("=NOT(E2)", "FALSE"),
    ('=IFS(B2>5,"a",B2>1,"b")', "b"), ('=CHOOSE(2,"x","y")', "y"),
    ('=CONCATENATE(A2," ",A3)', "Apple Pear"), ('=TEXTJOIN(", ",TRUE,A2:A5)', "Apple, Pear, Plum, Fig"),
    ("=LEN(A2)", "5"), ("=UPPER(A2)", "APPLE"), ('=TRIM("  a   b ")', "a b"), ("=LEFT(A2,3)", "App"),
    ("=RIGHT(A2,2)", "le"), ("=MID(A2,2,3)", "ppl"), ('=SUBSTITUTE("a-b-c","-","+")', "a+b+c"),
    ('=FIND("p",A2)', "2"), ('=SEARCH("P",A2)', "2"), ('=VALUE("12.5")', "12.5"),
    ('=TEXT(1234.5,"#,##0.00")', "1,234.50"), ('=TEXT(D2,"dd mmm yyyy")', "05 Mar 2024"),
    ("=DATE(2024,1,31)+1", "2024-02-01"), ("=YEAR(D2)", "2024"), ("=MONTH(D2)", "3"), ("=DAY(D3)", "10"),
    ("=D3-D2", "5"), ("=DATE(2024,14,1)", "2025-02-01"),
    ('=SUMIF(B2:B5,">3")', "15"), ('=SUMIF(A2:A5,"P*",B2:B5)', "5"), ('=COUNTIF(A2:A5,"p*")', "2"),
    ('=COUNTIF(B2:B5,"")', "1"), ('=AVERAGEIF(B2:B5,"<>3")', "7.5"),
    ('=SUMIFS(B2:B5,A2:A5,"<>Fig",C2:C5,">1")', "8"), ('=COUNTIFS(B2:B5,">=3",B2:B5,"<10")', "2"),
    ('=MAXIFS(B2:B5,A2:A5,"P*")', "5"),
    ('=VLOOKUP("pear",A2:C5,3,FALSE)', "2"), ('=VLOOKUP("zzz",A2:C5,2,FALSE)', "#N/A"),
    ("=VLOOKUP(4,B2:C3,2)", "1.5"), ('=HLOOKUP("Qty",A1:C3,2,FALSE)', "3"), ('=MATCH("Plum",A1:A5,0)', "4"),
    ("=INDEX(A1:C5,3,1)", "Pear"), ('=INDEX(B2:B5,MATCH("Fig",A2:A5,0))', "10"),
    ('=XLOOKUP("Plum",A2:A5,C2:C5)', "0.5"), ('=XLOOKUP("Kiwi",A2:A5,C2:C5,"none")', "none"),
    ("=ISBLANK(B4)", "TRUE"), ("=ISNUMBER(B2)", "TRUE"), ("=ISTEXT(A2)", "TRUE"), ("=ISERROR(C5*1)", "TRUE"),
    ("=ISNA(NA())", "TRUE"), ("=ROWS(A1:A5)", "5"),
    ("='My Sheet'!B2+1", "101"), ("=SUM('My Sheet'!B:B)", "300"), ("=Nope!A1", "#REF!"),
    ("=FOO(1)", "#NAME?"), ("=1+", "#NAME?"), ("=SUM(1,,2)", "3"), ('=_xlfn.CONCAT("a","b")', "ab"),
    ("=LARGE(B2:B5,2)", "5"), ("=SMALL(B2:B5,1)", "3"), ("=MEDIAN(B2:B5)", "5"),
])
def test_formula(ev, formula, want):
    assert ev(formula) == want


def test_circular_reference():
    s = Sheet([["=A2"], ["=A1"]])
    eng = Engine()
    eng.set_book({"S": s})
    assert eng.display(s, 0, 0) == "#CIRC!"


def test_long_chain_does_not_overflow():
    s = Sheet([["1"]] + [[f"=A{i}+1"] for i in range(1, 5000)])
    eng = Engine()
    eng.set_book({"C": s})
    t = time.time()
    assert eng.display(s, 4999, 0) == "5000"
    assert time.time() - t < 5


def test_unknown_function_falls_back_to_excel_result():
    s = Sheet([["=WEIRDFUNC(1)"]])
    s.cache = {(0, 0): "42"}
    eng = Engine()
    eng.set_book({"S": s})
    assert eng.display(s, 0, 0) == "42"


@pytest.mark.parametrize("formula, d, want", [
    ("=A1+B$2+$C3+$D$4", (1, 1), "=B2+C$2+$C4+$D$4"),
    ("=SUM(A1:A3)", (2, 0), "=SUM(A3:A5)"),
    ("=A1", (-1, 0), "=#REF!"),
    ("='My Sheet'!A1*2", (0, 2), "='My Sheet'!C1*2"),
    ('="A1"&A1', (1, 0), '="A1"&A2'),
    ("=SUM(B:B)", (0, 1), "=SUM(C:C)"),
])
def test_shift_on_paste(formula, d, want):
    assert shift(formula, *d) == want


@pytest.mark.parametrize("formula, home, target, axis, at, n, want", [
    ("=A5+B2", "S", "S", "row", 2, 1, "=A6+B2"),
    ("=SUM(A2:A5)", "S", "S", "row", 3, 2, "=SUM(A2:A7)"),
    ("=SUM(A2:A5)", "S", "S", "row", 1, 1, "=SUM(A3:A6)"),
    ("=A3", "S", "S", "row", 2, -1, "=#REF!"),
    ("=SUM(A2:A5)", "S", "S", "row", 2, -2, "=SUM(A2:A3)"),
    ("=SUM(A3:A4)", "S", "S", "row", 2, -2, "=SUM(#REF!)"),
    ("=SUM(A2:A5)", "S", "S", "row", 0, -2, "=SUM(A1:A3)"),
    ("=$C$3*D1", "S", "S", "col", 1, 1, "=$D$3*E1"),
    ("=SUM($A$2:$A$5)", "S", "S", "row", 3, 1, "=SUM($A$2:$A$6)"),
    ("=Other!A5", "S", "Other", "row", 0, 1, "=Other!A6"),
    ("=Other!A5+A5", "S", "S", "row", 0, 1, "=Other!A5+A6"),
    ("=SUM(B:D)", "S", "S", "col", 2, -1, "=SUM(B:C)"),
    ("=SUM(B:B)", "S", "S", "row", 0, 3, "=SUM(B:B)"),
    ("=SUM(2:4)", "S", "S", "row", 0, 1, "=SUM(3:5)"),
    ("='My S'!B2", "X", "My S", "col", 0, 1, "='My S'!C2"),
    ('="A5"&A5', "S", "S", "row", 0, 1, '="A5"&A6'),
])
def test_adjust_on_insert_delete(formula, home, target, axis, at, n, want):
    assert adjust(formula, home, target, axis, at, n) == want
