"""Scripts that manage the cells: Python inside the app, any program through CSV, --run at start."""
import sys

import openpyxl

import scripting
import workbook_tools as wt
import workbooks
from conftest import open_in


def run(app, code, **kw):
    return scripting.run(app, code, **kw)


def test_read_write_values_and_formulas(app):
    g = app.table
    app.set_cells([(0, 0, "3"), (1, 0, "4")])
    out = []
    ok, msg = run(app, '''
sheet["B1"] = sheet["A1"] * 10
sheet["A3"] = "=A1+A2"
print("sum is", sheet["A3"])            # a formula written by the script is calculated right away
sheet["C1"] = [[1, 2, 3], [4, 5, 6]]    # rows spread from one cell
sheet["D5:E6"] = "x"                    # one value fills a block
sheet.cell(8, 1, True)
''', out=out.append)
    assert ok and "changed" in msg and out == ["sum is 7\n"]
    assert g.text(0, 1) == "30" and g.text(2, 0) == "=A1+A2" and g.value(2, 0) == "7"
    assert [g.text(1, c) for c in (2, 3, 4)] == ["4", "5", "6"] and g.text(5, 4) == "x" and g.text(7, 0) == "TRUE"
    app.undo()                                                     # the whole run is one undo step
    assert g.text(0, 1) == "" and g.text(2, 0) == "" and g.text(5, 4) == ""


def test_reading_blocks_and_types(app):
    app.set_cells([(0, 0, "Name"), (0, 1, "Qty"), (1, 0, "a"), (1, 1, "2.5"), (2, 0, "b"), (2, 1, "")])
    out = []
    ok, _ = run(app, 'print(sheet["A1:B3"], sheet.last_row, sheet.last_col, sheet["B2"] + 1, sheet["B3"])',
                out=out.append)
    assert ok and out == ["[['Name', 'Qty'], ['a', 2.5], ['b', None]] 3 2 3.5 None\n"]


def test_error_and_stop_change_nothing(app):
    g = app.table
    app.set_cells([(0, 0, "keep")])
    ok, msg = run(app, 'sheet["A1"] = "gone"\nsheet["B1"] = 1\nraise ValueError("oops")')
    assert not ok and "line 3" in msg and "oops" in msg and "Nothing was changed" in msg
    assert g.text(0, 0) == "keep" and g.text(0, 1) == ""
    calls = [0]

    def stop():
        calls[0] += 1
        return calls[0] > 2
    ok, msg = run(app, 'sheet["A1"] = "gone"\nwhile True:\n    pass', stop=stop)
    assert not ok and msg.startswith("Stopped") and g.text(0, 0) == "keep"
    ok, msg = run(app, 'try:\n    while True:\n        pass\nexcept Exception:\n    pass', timeout=0.2)
    assert not ok and msg.startswith("Stopped")                    # a script can't swallow the stop


def test_formats(app):
    g = app.table
    ok, _ = run(app, 'sheet["A1"] = "T"\nsheet.format("A1:B1", bold=True, bg="#ffff00", numfmt="0.00")')
    assert ok and g.sheet.styles[(0, 0)] == {"bold": True, "bg": "#ffff00", "numfmt": "0.00"}
    ok, msg = run(app, 'sheet.format("A1", colour="red")')
    assert not ok and "Unknown format" in msg
    ok, _ = run(app, 'sheet.format("A1", bold=False)')
    assert ok and "bold" not in g.sheet.styles[(0, 0)]


def test_book_sheets_and_protection(app):
    app.add_sheet()
    other = app.sheet
    names = list(app.sheets)
    app.show(names[0])
    ok, _ = run(app, f'book["{other}"]["A1"] = "there"\nprint([s.name for s in book.sheets])')
    assert ok and app.sheets[other].rows[0][0] == "there"
    ok, msg = run(app, 'book["Nope"]')
    assert not ok and "No sheet named" in msg
    app.show(other)
    wt.protect_sheet(app, "")
    ok, msg = run(app, 'sheet["A1"] = "x"')
    assert not ok and "protected" in msg and app.sheets[other].rows[0][0] == "there"
    wt.unprotect_sheet(app, "")


def test_selection_and_examples(app):
    g = app.table
    app.set_cells([(1, 1, "2"), (2, 1, "5"), (1, 2, "abc")])
    g.select(1, 1, 2, 2)
    ok, _ = run(app, scripting.SNIPPETS["Double the numbers in the selection"])
    assert ok and g.text(1, 1) == "4" and g.text(2, 1) == "10" and g.text(1, 2) == "abc"
    ok, _ = run(app, scripting.SNIPPETS["Upper-case the selected text"])
    assert ok and g.text(1, 2) == "ABC"
    app.set_cells([(4, 0, "-3")])
    ok, _ = run(app, scripting.SNIPPETS["Highlight negative numbers in red"])
    assert ok and g.sheet.styles[(4, 0)]["fg"] == "#d93025"
    out = []
    assert run(app, scripting.SNIPPETS["List every sheet and its size"], out=out.append)[0] and out
    assert run(app, 'print(selection.address, selection.row, selection.col, ref(2, 3))', out=out.append)[0]
    assert out[-1] == "B2:C3 2 2 C2\n"


def test_script_can_save(app, tmp_path):
    path = workbooks.plain(tmp_path / "s.xlsx")
    open_in(app, path)
    ok, _ = run(app, 'sheet["F1"] = "from a script"\nbook.save()')
    assert ok and openpyxl.load_workbook(path).active["F1"].value == "from a script"


def test_script_window(app, tmp_path):
    w = scripting.open_window(app)
    w.load_text('sheet["A1"] = 1\nprint("hello")')
    assert w.run() is True
    out = w.output.get("1.0", "end")
    assert "hello" in out and "Done: 1 cell changed" in out and app.table.text(0, 0) == "1"
    w.snippet.set("Number the rows (1, 2, 3…) in column A")
    w.snippet.event_generate("<<ComboboxSelected>>")
    assert "last_row" in w.code.get("1.0", "end")
    p = w.save(str(tmp_path / "x.py"))
    w.load_text("")
    assert w.open(p) == p and "last_row" in w.code.get("1.0", "end")
    w.load_text("1/0")
    assert w.run() is False and "ZeroDivisionError" in w.output.get("1.0", "end")
    w.win.destroy()


def test_run_command_any_program(app):
    g = app.table
    app.set_cells([(0, 0, "a"), (0, 1, "b"), (1, 0, "c"), (1, 1, "d")])
    g.select(0, 0, 1, 1)
    upper = f'"{sys.executable}" -c "import sys; sys.stdout.write(sys.stdin.read().upper())"'
    ok, msg = scripting.run_command(app, upper)
    assert ok and [g.text(0, 0), g.text(1, 1)] == ["A", "D"]
    app.undo()
    assert g.text(0, 0) == "a"
    grow = f'"{sys.executable}" -c "print(\'x,y,z\'); print(\'1,2,3\'); print(\'4,5,6\')"'
    ok, _ = scripting.run_command(app, grow)                        # more than the selection: it spreads
    assert ok and g.text(2, 2) == "6"
    ok, msg = scripting.run_command(app, f'"{sys.executable}" -c "import sys; sys.exit(3)"')
    assert not ok and "exit code 3" in msg
    ok, msg = scripting.run_command(app, f'"{sys.executable}" -c "pass"')
    assert not ok and "printed nothing" in msg


def test_run_at_start(app, tmp_path):
    script = tmp_path / "job.py"
    script.write_text('sheet["A1"] = "auto"\n', encoding="utf-8")
    ok, msg = scripting.run_at_start(app, str(script))
    assert ok and app.table.text(0, 0) == "auto" and "OK" in (tmp_path / "job.log").read_text()
    ok, msg = scripting.run_at_start(app, str(tmp_path / "missing.py"))
    assert not ok and "Can't read" in msg


def guide_examples(section_names):
    """Fenced code blocks under the given headings of the script guide."""
    import help_ui
    blocks, section, cur = [], None, None
    for line in help_ui.guide_text("SCRIPT_GUIDE.md").split("\n"):
        if line.startswith("```"):
            if cur is None:
                cur = []
            else:
                if section in section_names:
                    blocks.append("\n".join(cur))
                cur = None
        elif cur is not None:
            cur.append(line)
        elif line.startswith("# "):
            section = line[2:].strip()
    return blocks


def test_help_menu_shows_script_guide(app):
    import help_ui
    w = help_ui.show(app, "scripts")
    text = w.text.get("1.0", "end")
    assert w.title() == "Script guide" and "Writing cells" in text and "```" not in text
    assert w.text.tag_ranges("code")                               # examples shaded
    w.destroy()
    help_ = next(app.app_menu.nametowidget(app.app_menu.entrycget(i, "menu"))
                 for i in range(app.app_menu.index("end") + 1)
                 if app.app_menu.type(i) == "cascade" and app.app_menu.entrycget(i, "label") == "Help")
    labels = [help_.entrycget(i, "label") for i in range(help_.index("end") + 1) if help_.type(i) == "command"]
    assert "Script guide" in labels


def test_guide_examples_really_run(app):
    """Every example in 'Your first script' and 'More examples' works as printed."""
    g = app.table
    blocks = guide_examples({"Your first script", "More examples"})
    assert len(blocks) >= 5
    for code in blocks:
        app.set_cells([(0, 0, "Name"), (0, 1, "Qty"), (0, 2, "Price"),
                       (1, 0, "Lima, Ana"), (1, 1, "3"), (1, 2, "-2.5"),
                       (2, 0, "Ortiz, Ben"), (2, 1, "4"), (2, 2, "10")])
        g.select(1, 1, 2, 2)
        ok, msg = scripting.run(app, code)
        assert ok, (code, msg)
        app.undo()
