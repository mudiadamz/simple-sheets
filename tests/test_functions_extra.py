"""Finance, statistics, math, date/time, text and reference functions (functions_extra.py)."""
import pytest

from formula import Engine
from sheet_io import Sheet


@pytest.fixture
def ev():
    data = Sheet([
        ["n", "flow", "d"],
        ["2", "-100", "2024-01-31"],
        ["4", "30", "2024-03-01"],
        ["4", "40", ""],
        ["5", "50", ""],
        ["7", "", ""],
    ])
    eng = Engine()
    eng.set_book({"Data": data, "Two": Sheet([["x"], ["9"]])})

    def run(formula):
        data.rows.append([formula])
        eng.clear()
        try:
            return eng.display(data, len(data.rows) - 1, 0)
        finally:
            data.rows.pop()
    return run


@pytest.mark.parametrize("formula, want", [
    # finance
    ("=ROUND(PMT(0.05/12,60,10000),2)", "-188.71"), ("=PMT(0,10,100)", "-10"),
    ("=ROUND(FV(0.06/12,10,-200,-500,1),2)", "2581.4"), ("=ROUND(PV(0.08/12,240,500),2)", "-59777.15"),
    ("=ROUND(NPER(0.01,-100,1000),4)", "10.5886"), ("=ROUND(RATE(60,-188.71,10000)*12,4)", "0.05"),
    ("=ROUND(NPV(0.1,-10000,3000,4200,6800),2)", "1188.44"), ("=ROUND(IRR(B2:B5),4)", "0.089"),
    ("=IRR(A2:A3)", "#NUM!"),
    # statistics
    ("=ROUND(STDEV(A2:A6),4)", "1.8166"), ("=ROUND(STDEV.P(A2:A6),4)", "1.6248"), ("=VAR(A2:A6)", "3.3"),
    ("=VARP(A2:A6)", "2.64"), ("=MODE(A2:A6)", "4"), ("=MODE(1,2,3)", "#N/A"), ("=PERCENTILE(A2:A6,0.5)", "4"),
    ("=QUARTILE(A2:A6,1)", "4"), ("=QUARTILE(A2:A6,4)", "7"), ("=RANK(5,A2:A6)", "2"), ("=RANK(5,A2:A6,1)", "4"),
    ("=RANK(9,A2:A6)", "#N/A"), ("=COUNTUNIQUE(A2:A6)", "4"), ("=STDEV(1)", "#DIV/0!"),
    # math
    ("=SUMPRODUCT(A2:A3,B2:B3)", "-80"), ("=SUMPRODUCT(A2:A3,B2:B4)", "#VALUE!"), ("=SUMSQ(3,4)", "25"),
    ("=CEILING(2.1,0.5)", "2.5"), ("=FLOOR(2.9)", "2"), ("=MROUND(10,3)", "9"), ("=MROUND(-10,3)", "#NUM!"),
    ("=EVEN(1.5)", "2"), ("=EVEN(-3)", "-4"), ("=ODD(2)", "3"), ("=ODD(-1.5)", "-3"), ("=FACT(5)", "120"),
    ("=FACT(-1)", "#NUM!"), ("=GCD(24,36)", "12"), ("=LCM(4,6)", "12"), ("=ROUND(SIN(PI()/2),6)", "1"),
    ("=ROUND(COS(0),6)", "1"), ("=DEGREES(PI())", "180"), ("=ROUND(RADIANS(180),5)", "3.14159"),
    ("=ROUND(ATAN2(1,1),4)", "0.7854"), ("=ASIN(2)", "#NUM!"), ("=ROUND(ACOS(0),4)", "1.5708"),
    # dates & times
    ('=TEXT(EDATE(C2,1),"yyyy-mm-dd")', "2024-02-29"), ('=TEXT(EOMONTH(C2,1),"yyyy-mm-dd")', "2024-02-29"),
    ('=TEXT(EOMONTH(C3,-1),"yyyy-mm-dd")', "2024-02-29"), ("=NETWORKDAYS(C2,C3)", "23"),
    ("=NETWORKDAYS(C2,C3,C3)", "22"), ('=TEXT(WORKDAY(C2,1),"yyyy-mm-dd")', "2024-02-01"),
    ('=TEXT(WORKDAY(C3,-1),"yyyy-mm-dd")', "2024-02-29"), ('=DATEDIF(C2,C3,"D")', "30"),
    ('=DATEDIF(C2,C3,"M")', "1"), ('=DATEDIF("2020-05-10","2024-03-01","Y")', "3"),
    ('=DATEDIF("2020-05-10","2024-03-01","YM")', "9"), ('=DATEDIF(C3,C2,"D")', "#NUM!"),
    ("=HOUR(0.75)", "18"), ("=MINUTE(TIME(1,30,15))", "30"), ("=SECOND(TIME(1,30,15))", "15"),
    ('=TEXT(DATEVALUE("2024-03-01"),"dd/mm/yyyy")', "01/03/2024"),
    # text
    ("=CHAR(65)", "A"), ('=CODE("a")', "97"), ('=CLEAN("a"&CHAR(9)&"b")', "ab"), ("=FIXED(1234.567,1)", "1,234.6"),
    ("=FIXED(1234.567,0,TRUE)", "1235"), ("=ISEVEN(4)", "TRUE"), ("=ISODD(4)", "FALSE"),
    # references
    ("=ROW()", "7"), ("=COLUMN()", "1"), ("=ROW(B3)", "3"), ("=COLUMN(C1:D4)", "3"), ("=ROW(Two!A2)", "2"),
    ("=SUM(OFFSET(A1,1,0,3,1))", "10"), ("=OFFSET(A1,2,1)", "30"), ("=OFFSET(A1,-1,0)", "#REF!"),
    ('=INDIRECT("A3")', "4"), ('=SUM(INDIRECT("A2:A4"))', "10"), ('=INDIRECT("Two!A2")', "9"),
    ('=INDIRECT("nope!!")', "#REF!"), ("=ADDRESS(2,3)", "$C$2"), ("=ADDRESS(2,3,4)", "C2"),
    ('=ADDRESS(1,1,1,TRUE,"My Sheet")', "'My Sheet'!$A$1"),
])
def test_function(ev, formula, want):
    assert ev(formula) == want
