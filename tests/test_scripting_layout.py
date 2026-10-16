"""Scripts (and so the AI) can change the sheet's layout too: freeze, merge, sizes, hiding, insert / delete,
sort, filter, drop-downs, conditional formats, borders, comments, links, tab colour, new sheets. One undo
step each run; a failing run changes nothing - structural changes included."""
import openpyxl

import scripting
import workbooks
from conftest import open_in


def run(app, code):
    ok, msg = scripting.run(app, code)
    assert ok, msg
    return msg


def table(app):
    rows = [["Name", "Qty", "Region"], ["b", "5", "East"], ["a", "50", "West"], ["c", "15", "East"]]
    app.set_cells([(r, c, v) for r, row in enumerate(rows) for c, v in enumerate(row)])
    return app.table


def test_freeze_merge_sizes_hiding(app):
    g = table(app)
    run(app, '''
sheet.freeze(1)
sheet.merge("A6:C6")
sheet.width("B:C", 20)
sheet.height(1, 30)
sheet.hide_rows("4")
sheet.hide_cols("C")
''')
    s = g.sheet
    assert s.freeze == (1, 0) and (5, 0, 5, 2) in s.merges and s.widths[1] == s.widths[2] == 20.0
    assert s.heights[0] == 30.0 and 3 in s.hidden_rows and 2 in s.hidden_cols
    assert g.fr == 1                                                   # the grid shows it frozen right away
    app.undo()                                                         # the whole run in one step
    assert s.freeze == (0, 0) and not s.merges and 1 not in s.widths and not s.hidden_rows and not s.hidden_cols
    run(app, 'sheet.freeze("B2")\nsheet.autofit()')
    assert s.freeze == (1, 1) and s.widths[0] >= 6
    run(app, 'sheet.unfreeze()\nsheet.show_rows("4")')
    assert s.freeze == (0, 0)


def test_insert_delete_rows_and_columns(app):
    g = table(app)
    app.set_cells([(0, 4, "=SUM(B2:B4)")])
    run(app, 'sheet.insert_rows(2, 2)\nsheet["A2"] = "new"')
    assert g.text(1, 0) == "new" and g.text(3, 0) == "b" and g.text(0, 4) == "=SUM(B4:B6)"   # formulas follow
    run(app, 'sheet.delete_rows(2, 2)\nsheet.insert_cols("B")')
    assert g.text(1, 0) == "b" and g.text(0, 2) == "Qty" and g.text(0, 5) == "=SUM(C2:C4)"
    app.undo()
    assert g.text(0, 1) == "Qty" and g.text(1, 0) == "new"                                 # undo: one step
    app.undo()
    assert g.text(1, 0) == "b" and g.text(0, 4) == "=SUM(B2:B4)"


def test_sort(app):
    g = table(app)
    g.select(1, 0, 1, 0)
    app.format_cells(lambda st: {**st, "bold": True})                 # "b" row is bold
    run(app, 'sheet.sort("A1:C4", by="B", descending=True)')
    assert [g.text(r, 0) for r in range(1, 4)] == ["a", "c", "b"]
    assert g.sheet.styles[(3, 0)]["bold"]                              # the format moved with its row
    run(app, 'sheet.sort("A1:C4", by=["C", "A"])')
    assert [g.text(r, 0) for r in range(1, 4)] == ["b", "c", "a"]


def test_filter_dropdown_rules_borders(app):
    g = table(app)
    run(app, '''
sheet.filter("A1:C4", column="C", values=["East"])
sheet.dropdown("C2:C10", ["East", "West"])
sheet.highlight("B2:B4", ">", 10, bg="#ffc7ce")
sheet.color_scale("B2:B4")
sheet.data_bar("B2:B4")
sheet.format("A1:C1", border="outline", border_style="medium", bold=True)
''')
    s = g.sheet
    assert s.filtered_rows == {2}                                      # West hidden
    assert s.validations[0]["f1"] == '"East,West"'
    assert [d["type"] for d in s.cond] == ["dataBar", "colorScale", "cellIs"]
    assert s.styles[(0, 0)]["border"] == {"top": ["medium", "#000000"], "bottom": ["medium", "#000000"],
                                          "left": ["medium", "#000000"]}
    assert "left" not in s.styles[(0, 1)]["border"] and "right" in s.styles[(0, 2)]["border"]
    assert app.rules.conditional(s, 2, 1)                              # 50 > 10: highlighted
    run(app, 'sheet.clear_rules()\nsheet.clear_filter()')
    assert not s.cond and not s.validations and s.filter is None and not s.filtered_rows


def test_comment_link_tab_and_new_sheet(app):
    table(app)
    first = app.sheet
    run(app, '''
sheet.comment("A1", "names here")
sheet.link("A2", "https://example.com")
sheet.tab_color("#ff0000")
s2 = book.add_sheet("Summary")
s2["A1"] = "=SUM(''' + first + '''!B2:B4)"
''')
    s = app.sheets[first]
    assert s.comments[(0, 0)][0] == "names here" and s.links[(1, 0)] == "https://example.com"
    assert s.props["tab"] == "#ff0000" and "Summary" in app.sheets
    assert app.sheets["Summary"].rows[0][0].startswith("=SUM(")
    app.undo()                                                         # one step: the sheet and the rest go
    assert "Summary" not in app.sheets and (0, 0) not in s.comments and "tab" not in s.props
    run(app, f'book.rename_sheet("{first}", "Data")\nsheet["D1"] = "x"')
    assert "Data" in app.sheets and app.sheets["Data"].rows[0][3] == "x"


def test_failed_run_undoes_structure_too(app):
    g = table(app)
    ok, msg = scripting.run(app, 'sheet.insert_rows(1, 3)\nbook.add_sheet("X")\nsheet.freeze(2)\n1/0')
    assert not ok and "Nothing was changed" in msg
    assert g.text(0, 0) == "Name" and "X" not in app.sheets and g.sheet.freeze == (0, 0)
    ok, msg = scripting.run(app, 'sheet.width("B", 30)\nsheet.merge("A1:B1")\nraise ValueError("x")')
    assert not ok and 1 not in g.sheet.widths and not g.sheet.merges


def test_errors_are_clear(app):
    table(app)
    for code, text in (('sheet.width("1x", 5)', "Not a column"), ('sheet.highlight("A1", "~", 1)', "unknown rule"),
                       ('sheet.format("A1", border="diagonal")', "border:"),
                       ('sheet.format("A1", border="all", border_style="wavy")', "border_style"),
                       ('book.rename_sheet("Nope", "X")', "No sheet")):
        ok, msg = scripting.run(app, code)
        assert not ok and text in msg, (code, msg)


def test_layout_saved(app, tmp_path):
    path = workbooks.plain(tmp_path / "l.xlsx")
    open_in(app, path)
    table(app)
    run(app, '''
sheet.freeze(1)
sheet.merge("E1:F1")
sheet.width("B", 25)
sheet.hide_cols("D")
sheet.dropdown("C2:C9", ["East", "West"])
sheet.highlight("B2:B9", ">", 10)
sheet.format("A1:C1", border="all")
book.save()
''')
    ws = openpyxl.load_workbook(path).active
    assert ws.freeze_panes == "A2" and "E1:F1" in [str(m) for m in ws.merged_cells.ranges]
    assert ws.column_dimensions["B"].width == 25 and ws.column_dimensions["D"].hidden
    assert ws.data_validations.dataValidation[0].formula1 == '"East,West"'
    assert len(ws.conditional_formatting) == 1 and ws["B1"].border.top.style == "thin"


def test_ai_is_told_about_layout_commands():
    import ai_assist
    for name in ("sheet.freeze", "sheet.merge", "sheet.sort", "sheet.filter", "sheet.dropdown", "sheet.highlight",
                 "insert_rows", "book.add_sheet", "border="):
        assert name in ai_assist.SYSTEM
    assert "never tell the user to do one of those by hand" in ai_assist.SYSTEM


def test_clear(app):
    g = table(app)
    g.select(0, 0, 0, 2)
    app.format_cells(lambda st: {**st, "bold": True})
    run(app, 'sheet.clear("A2:C3")')
    assert g.text(1, 0) == "" and g.text(2, 1) == "" and g.text(3, 0) == "c"
    run(app, "sheet.clear()")
    assert all(g.text(r, c) == "" for r in range(4) for c in range(3)) and g.sheet.styles[(0, 0)]["bold"]
    app.undo()
    run(app, "sheet.clear(formats=True)")
    assert not g.sheet.styles.get((0, 0))
