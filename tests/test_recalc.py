"""Only formulas that depend on an edited cell are calculated again."""
from formula import Engine
from sheet_io import Sheet


def counting(engine):
    calls = []
    real = engine.evaluate

    def spy(formula, sheet, spill=False):
        calls.append(formula)
        return real(formula, sheet, spill)
    engine.evaluate = spy
    return calls


def book():
    s = Sheet([["1", "=A1*2", "=B1+1", "5", "=D1*3", "=SUM(A:A)"], ["2"]])
    e = Engine()
    e.set_book({"S": s, "T": Sheet([["=S!A1+100"]])})
    return s, e, e.book["T"]


def show_all(e, s, t):
    return [e.display(s, 0, c) for c in (1, 2, 4, 5)] + [e.display(t, 0, 0)]


def test_only_dependents_recalculate():
    s, e, t = book()
    calls = counting(e)
    assert show_all(e, s, t) == ["2", "3", "15", "3", "101"]
    calls.clear()
    s.rows[0][3] = "6"                                     # D1: only E1 reads it
    e.invalidate({("S", 0, 3)})
    assert show_all(e, s, t) == ["2", "3", "18", "3", "101"]
    assert calls == ["=D1*3"]
    calls.clear()
    s.rows[0][0] = "10"                                    # A1: B1, C1 (via B1), F1 (range), T!A1
    e.invalidate({("S", 0, 0)})
    assert show_all(e, s, t) == ["20", "21", "18", "12", "110"]
    assert sorted(calls) == sorted(["=A1*2", "=B1+1", "=SUM(A:A)", "=S!A1+100"])


def test_new_cell_inside_a_whole_column_reference():
    s, e, t = book()
    show_all(e, s, t)
    s.rows.append(["7"])                                   # A3 didn't exist: SUM(A:A) still updates
    e.invalidate({("S", 2, 0)})
    assert e.display(s, 0, 5) == "10"


def test_volatile_always_recalculated():
    s = Sheet([["1", "=RAND()", "9"]])
    e = Engine()
    e.set_book({"S": s})
    first = e.display(s, 0, 1)
    s.rows[0][2] = "10"
    e.invalidate({("S", 0, 2)})                            # an unrelated cell changed: RAND() still redone
    assert e.display(s, 0, 1) != first


def test_app_edits_use_it(app):
    g = app.table
    app.set_cells([(0, 0, "1"), (0, 1, "=A1+1"), (5, 5, "=1+1")])
    assert g.value(0, 1) == "2" and g.value(5, 5) == "2"
    calls = counting(app.engine)
    app.set_cells([(0, 0, "5")])
    assert g.value(0, 1) == "6" and g.value(5, 5) == "2"
    assert calls == ["=A1+1"]                              # =1+1 wasn't calculated again
    app.format_cells(lambda st: {**st, "bold": True})      # formatting: nothing recalculated
    calls.clear()
    g.value(0, 1)
    assert calls == []
