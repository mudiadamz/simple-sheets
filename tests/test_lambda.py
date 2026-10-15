"""LET, LAMBDA (also named ones), MAP / BYROW / REDUCE / SCAN / MAKEARRAY, XMATCH, SWITCH, TEXTBEFORE / AFTER."""
import pytest

from formula import Engine
from sheet_io import Sheet


@pytest.fixture
def ev():
    s = Sheet([["3", "x"], ["5", "y"], ["2", "z"]])
    e = Engine()
    e.set_book({"S": s})
    e.set_names({"DOUBLE": "=LAMBDA(n, n*2)", "RATE": "S!$A$1"})

    def run(formula):
        s.rows.append([formula])
        e.clear()
        try:
            out = [e.display(s, len(s.rows) - 1, 0)]
            for r in range(len(s.rows), len(s.rows) + 3):      # spilled cells below
                v = e.spilled(s, r, 0)
                if v is None:
                    break
                out.append(v)
            return out[0] if len(out) == 1 else out
        finally:
            s.rows.pop()
    return run


@pytest.mark.parametrize("formula, want", [
    ("=LET(x, 2, y, x*3, x+y)", "8"),
    ("=LET(total, SUM(A1:A3), total/2)", "5"),
    ("=LAMBDA(a, b, a*b)(6, 7)", "42"),
    ("=DOUBLE(21)", "42"),
    ("=DOUBLE(RATE)", "6"),
    ("=LAMBDA(x, x)", "#CALC!"),
    ("=LAMBDA(x, x*2)(1, 2)", "#VALUE!"),
    ("=REDUCE(0, A1:A3, LAMBDA(acc, v, acc+v))", "10"),
    ("=SUM(MAP(A1:A3, LAMBDA(v, v*10)))", "100"),
    ("=MAP(A1:A3, LAMBDA(v, v+1))", ["4", "6", "3"]),
    ("=BYROW(A1:B2, LAMBDA(r, COUNTA(r)))", ["2", "2"]),
    ("=SCAN(0, A1:A3, LAMBDA(a, v, a+v))", ["3", "8", "10"]),
    ("=MAKEARRAY(2, 1, LAMBDA(r, c, r*10))", ["10", "20"]),
    ("=XMATCH(5, A1:A3)", "2"),
    ("=XMATCH(4, A1:A3, 1)", "2"),
    ("=XMATCH(4, A1:A3, -1)", "1"),
    ('=XMATCH("y*", B1:B3, 2)', "2"),
    ("=XMATCH(9, A1:A3)", "#N/A"),
    ('=SWITCH(A2, 3, "three", 5, "five", "other")', "five"),
    ('=SWITCH(9, 1, "one")', "#N/A"),
    ('=TEXTBEFORE("a-b-c", "-")', "a"),
    ('=TEXTAFTER("a-b-c", "-", 2)', "c"),
    ('=TEXTBEFORE("a-b-c", "-", -1)', "a-b"),
    ('=TEXTAFTER("abc", "x")', "#N/A"),
    ('=ISOWEEKNUM(DATE(2024,1,1))', "1"),
    ('=WEEKNUM(DATE(2024,1,7))', "2"),
])
def test_functions(ev, formula, want):
    assert ev(formula) == want
