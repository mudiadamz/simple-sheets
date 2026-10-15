"""Find & replace."""
import find_replace
from conftest import key


def setup_book(app):
    app.set_cells([(0, 0, "Apple"), (1, 0, "apple pie"), (2, 0, "Pear"), (3, 1, "=A1&\"!\"")])
    app.add_sheet()
    app.set_cells([(5, 2, "APPLE")])
    app.show("Sheet1")
    app.tabs.select("Sheet1")
    app.table.goto(0, 0)


def test_find_all_options(app):
    setup_book(app)
    hits = lambda **k: [(n, r, c) for n, r, c in find_replace.find_all(app, "apple", **k)]
    assert hits() == [("Sheet1", 0, 0), ("Sheet1", 1, 0), ("Sheet1", 3, 1)]       # B4 shows "Apple!"
    assert hits(match_case=True) == [("Sheet1", 1, 0)]
    assert hits(whole=True) == [("Sheet1", 0, 0)]
    assert hits(in_formulas=True) == [("Sheet1", 0, 0), ("Sheet1", 1, 0)]          # "=A1&..." has no "apple"
    assert hits(workbook=True)[-1] == ("Sheet2", 5, 2)


def test_dialog_find_next_wraps_and_switches_sheet(app):
    setup_book(app)
    key(app, "<Control-f>")
    d = app.find_dialog
    d.find.delete(0, "end")
    d.find.insert(0, "apple")
    d.workbook.set(True)
    d.reset()
    found = [d.next() for _ in range(5)]
    assert [f[:3] for f in found[:4]] == [("Sheet1", 1, 0), ("Sheet1", 3, 1), ("Sheet2", 5, 2), ("Sheet1", 0, 0)]
    assert app.sheet == "Sheet1" and app.table.cur == (1, 0)      # wrapped round
    assert d.status["text"] == "2 of 4"
    d.destroy()


def test_replace_one_and_all_with_undo(app):
    setup_book(app)
    d = find_replace.FindDialog(app, replace=True)
    d.find.insert(0, "apple")
    d.repl.insert(0, "Mango")
    d.next()                                                    # A2
    d.replace_one()
    assert app.table.text(1, 0) == "Mango pie"
    d.workbook.set(True)
    d.reset()
    n = d.replace_all()
    assert n == 2 and app.table.text(0, 0) == "Mango" and app.sheets["Sheet2"].rows[5][2] == "Mango"
    assert app.table.text(3, 1) == '=A1&"!"'                    # formula untouched (searching values)
    app.undo()                                                  # replace all = one step
    assert app.table.text(0, 0) == "Apple" and app.sheets["Sheet2"].rows[5][2] == "APPLE"
    d.destroy()


def test_replace_in_formulas(app):
    setup_book(app)
    d = find_replace.FindDialog(app, replace=True)
    d.find.insert(0, "A1")
    d.repl.insert(0, "A3")
    d.formulas.set(True)
    d.replace_all()
    assert app.table.text(3, 1) == '=A3&"!"' and app.table.value(3, 1) == "Pear!"
    d.destroy()


def test_menu_entries(app):
    labels = [app.app_menu.entrycget(i, "label") for i in range(app.app_menu.index("end") + 1)
              if app.app_menu.type(i) != "separator"]
    assert "Find…" in labels and "Replace…" in labels


def test_find_all_list_across_sheets(app):
    import find_replace
    app.set_cells([(0, 0, "apple pie"), (3, 1, "=\"apple\"&\"s\"")])
    app.add_sheet()
    app.set_cells([(2, 2, "Apple tree")])
    find_replace.open_dialog(app)
    d = app.find_dialog
    d.find.delete(0, "end")
    d.find.insert(0, "apple")
    d.workbook.set(True)
    hits = d.find_all_list()
    assert len(hits) == 3 and d.results_shown
    rows = [d.results.item(i, "values") for i in d.results.get_children()]
    assert ("Sheet2", "C3", "Apple tree") in [tuple(r) for r in rows]
    assert any(r[2] == "apples" for r in rows)          # formulas are searched by what they show
    d.results.selection_set("2")
    d.go_result()
    assert (app.sheet, app.table.cur) == (d.hits[2][0], d.hits[2][1:])
    d.destroy()
