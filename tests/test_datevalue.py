"""Dates and times written as text: DATEVALUE, TIMEVALUE, and "15 Oct 2026" + 1 (formatting.parse_date_text)."""
import datetime as dt

import pytest

import formatting
from formula import Engine
from sheet_io import Sheet

TODAY = dt.date(2026, 10, 16)


@pytest.fixture
def ev():
    data = Sheet([["15 Oct 2026", "Oktober 2026"]])
    eng = Engine()
    eng.set_book({"Data": data})

    def run(formula):
        data.rows.append([formula])
        eng.clear()
        try:
            return eng.display(data, len(data.rows) - 1, 0)
        finally:
            data.rows.pop()
    return run


@pytest.fixture(autouse=True)
def month_first(monkeypatch):
    monkeypatch.setattr(formatting, "DATE_ORDER", "mdy")      # like Excel in English (US); tests can switch


@pytest.mark.parametrize("text, want", [
    ("2026-10-15", (2026, 10, 15)), ("15 Oct 2026", (2026, 10, 15)), ("15-Oct-2026", (2026, 10, 15)),
    ("15-Oct-26", (2026, 10, 15)), ("Oct 15, 2026", (2026, 10, 15)), ("October 15 2026", (2026, 10, 15)),
    ("2026 Oct 15", (2026, 10, 15)), ("Thu, 15 Oct 2026", (2026, 10, 15)), ("15. Okt 2026", (2026, 10, 15)),
    ("17 Agustus 1945", (1945, 8, 17)), ("1 Mei 2026", (2026, 5, 1)), ("Oct 2026", (2026, 10, 1)),
    ("October 1999", (1999, 10, 1)), ("15 Oct", (2026, 10, 15)), ("Sept 3", (2026, 9, 3)),
    ("10/15/2026", (2026, 10, 15)), ("15/10/2026", (2026, 10, 15)), ("2026/10/15", (2026, 10, 15)),
    ("3/4/2026", (2026, 3, 4)), ("10.15.26", (2026, 10, 15)), ("1/2/99", (1999, 1, 2)), ("10/15", (2026, 10, 15)),
])
def test_dates(text, want):
    assert formatting.parse_date_text(text, today=TODAY) == dt.datetime(*want)


def test_day_first_windows_setting(monkeypatch):
    monkeypatch.setattr(formatting, "DATE_ORDER", "dmy")
    assert formatting.parse_date_text("3/4/2026") == dt.datetime(2026, 4, 3)
    assert formatting.parse_date_text("10/15/2026") == dt.datetime(2026, 10, 15)   # only one way it makes sense


@pytest.mark.parametrize("text, want", [
    ("15 Oct 2026 14:30", (2026, 10, 15, 14, 30, 0)), ("Oct 15, 2026 2:30 PM", (2026, 10, 15, 14, 30, 0)),
    ("2:30 PM", (1899, 12, 30, 14, 30, 0)), ("14:30:15", (1899, 12, 30, 14, 30, 15)), ("9am", (1899, 12, 30, 9, 0, 0)),
    ("12:05 AM", (1899, 12, 30, 0, 5, 0)),
])
def test_times(text, want):
    assert formatting.parse_date_text(text, today=TODAY) == dt.datetime(*want)


@pytest.mark.parametrize("text", ["", "Oct", "May", "hello", "31 Feb 2026", "13/13/2026", "15 Oct 2026 25:00",
                                  "1 2 3 4", "Oct Nov 2026", "1.5", "13:00 PM", "abc 2026"])
def test_not_dates(text):
    assert formatting.parse_date_text(text, today=TODAY) is None


@pytest.mark.parametrize("formula, want", [
    ('=TEXT(DATEVALUE("15 Oct 2026"),"yyyy-mm-dd")', "2026-10-15"),
    ('=TEXT(DATEVALUE("October 15, 2026 6:00 PM"),"yyyy-mm-dd hh:mm")', "2026-10-15 00:00"),   # time dropped
    ('=MONTH(DATEVALUE("1 " & B1))', "10"),                       # "Oktober 2026" in a cell: the month number
    ('=DATEVALUE("2:30 PM")', "#VALUE!"), ('=DATEVALUE("hello")', "#VALUE!"),
    ('=TEXT(TIMEVALUE("2:30 PM"),"0.0000")', "0.6042"), ('=TEXT(TIMEVALUE("15 Oct 2026 18:45"),"hh:mm")', "18:45"),
    ('=TIMEVALUE("nope")', "#VALUE!"),
    ('=TEXT(A1+1,"yyyy-mm-dd")', "2026-10-16"),                    # text dates work in arithmetic, like Excel
    ('=YEAR("Oct 15, 2026")', "2026"), ('=VALUE("15 Oct 2026")', "46310"),
    ('="abc"+1', "#VALUE!"),
])
def test_in_formulas(ev, formula, want):
    assert ev(formula) == want
