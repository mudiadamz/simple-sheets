"""Moving cells (drag the selection's edge) updates formulas that point at them, like Excel."""
from formula import move_refs


def test_move_refs_text():
    assert move_refs("=A1+B2", "S", "S", (0, 0, 1, 1), 5, 2) == "=C6+D7"
    assert move_refs("=$A$1*2", "S", "S", (0, 0, 0, 0), 1, 1) == "=$B$2*2"      # absolute too
    assert move_refs("=SUM(A1:A2)", "S", "S", (0, 0, 1, 0), 3, 0) == "=SUM(A4:A5)"
    assert move_refs("=SUM(A1:A9)", "S", "S", (0, 0, 1, 0), 3, 0) == "=SUM(A1:A9)"  # only partly inside
    assert move_refs("=S!A1", "Other", "S", (0, 0, 0, 0), 0, 1) == "=S!B1"
    assert move_refs("=A1", "Other", "S", (0, 0, 0, 0), 0, 1) == "=A1"           # A1 of another sheet
    assert move_refs("plain", "S", "S", (0, 0, 0, 0), 1, 1) == "plain"


def test_moving_block_updates_references(app):
    g = app.table
    app.set_cells([(0, 0, "5"), (1, 0, "=A1*2"), (0, 3, "=A1+A2"), (0, 4, "=SUM(A1:A2)")])
    app.add_sheet()
    other = app.sheet
    app.set_cells([(0, 0, "=Sheet1!A1")])
    app.show("Sheet1")
    app.tabs.select("Sheet1")
    assert app.move_block((0, 0, 1, 0), 4, 1)           # A1:A2 -> B5:B6
    assert g.text(4, 1) == "5" and g.text(5, 1) == "=B5*2"          # the moved formula follows its partner
    assert g.text(0, 3) == "=B5+B6" and g.text(0, 4) == "=SUM(B5:B6)"
    assert app.sheets[other].rows[0][0] == "=Sheet1!B5"
    assert g.value(0, 3) == "15"
    app.undo()
    assert g.text(0, 3) == "=A1+A2" and app.sheets[other].rows[0][0] == "=Sheet1!A1" and g.text(0, 0) == "5"
