"""Formula autocomplete: function list while typing, argument hints, keys."""
from autocomplete import context
from conftest import key


def test_context():
    assert context("=SU", 3) == (("SU", 1), None)
    assert context("=A1+VLO", 7) == (("VLO", 4), None)
    assert context("=SUM(A1,", 8) == (None, ("SUM", 1))
    assert context('=IF(A1>2,"x,y",RO', 17) == (("RO", 15), ("IF", 2))
    assert context('="SU', 4) == (None, None)                 # inside a string
    assert context("=Sheet2!A", 9)[0] is None                  # after a sheet name: a reference, not a name
    assert context("SU", 2) == (None, None)                    # not a formula
    assert context("=ROUND(SUM(A1,A2),", 18) == (None, ("ROUND", 1))


def edit(app, text):
    g = app.table
    g.goto(0, 0)
    g.edit(text)
    app.update()
    e = g.editor
    e.autocomplete.refresh()
    return e, e.autocomplete


def test_list_and_tab_accepts(app):
    e, ac = edit(app, "=SU")
    assert ac.items[0] == "SUM" and "SUMPRODUCT" in ac.items
    assert ac.top.winfo_ismapped() or ac.top.state() != "withdrawn"
    ac.move(1)
    ac.move(-1)
    key(app, "<Tab>", e)
    assert e.get() == "=SUM(" and app.table.editor is e        # Tab picked the name, edit still open
    assert not ac.items and "SUM(number1" in ac.hint.get("1.0", "end")   # now the argument hint
    e.insert("end", "1,2)")
    key(app, "<Return>", e)                                    # list closed: Enter commits as usual
    assert app.table.value(0, 0) == "3"


def test_hint_marks_current_argument(app):
    e, ac = edit(app, "=PMT(0.1,")
    assert not ac.items
    ranges = ac.hint.tag_ranges("arg")
    assert ac.hint.get(ranges[0], ranges[1]) == "nper"


def test_escape_closes_list_only(app):
    e, ac = edit(app, "=AVER")
    assert ac.items
    key(app, "<Escape>", e)
    assert not ac.items and app.table.editor is e              # the edit stays open
    key(app, "<Escape>", e)
    assert app.table.editor is None


def test_named_ranges_and_enter(app):
    app.set_names({"prices": "Sheet1!$B$1:$B$3"})
    e, ac = edit(app, "=SUM(pri")
    assert ac.items == ["prices"]
    key(app, "<Return>", e)
    assert e.get() == "=SUM(prices" and app.table.editor is e


def test_formula_bar(app):
    app.table.goto(0, 0)
    v = app.value
    v.delete(0, "end")
    v.insert(0, "=MA")
    v.icursor("end")
    v.autocomplete.refresh()
    assert "MAX" in v.autocomplete.items
    v.autocomplete.accept()
    assert v.get().startswith("=MA") and v.get().endswith("(")       # the first suggestion went in


def test_no_list_for_plain_text(app):
    e, ac = edit(app, "Sum of")
    assert not ac.items
