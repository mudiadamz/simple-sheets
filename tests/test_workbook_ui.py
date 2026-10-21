"""☰ menu, new window, open recent, blank start, adding / renaming / deleting sheets."""
import os
from types import SimpleNamespace

import openpyxl
import pytest
from simplecore import recent

import workbooks
from conftest import key, open_in


def all_texts(w):
    out = []
    for child in w.winfo_children():
        try:
            out.append(str(child.cget("text")))
        except Exception:
            pass
        out += all_texts(child)
    return out


def menu_labels(menu):
    return [menu.entrycget(i, "label") for i in range(menu.index("end") + 1) if menu.type(i) != "separator"]


# ---------- menu / layout ----------

def test_hamburger_menu_replaces_toolbar(app):
    texts = all_texts(app)
    from simplecore import icons
    assert icons.text(app, "menu") in texts                     # ☰ (Windows: its own icon font)
    assert not {"Open…", "Save", "Save As…", "↶ Undo", "↷ Redo"} & set(texts)
    labels = menu_labels(app.app_menu)
    core = ["New", "Open…", "Open recent", "Save", "Save as…", "Undo", "Redo", "View", "Exit"]
    assert [x for x in labels if x in core] == core                   # in this order; other items may sit between
    assert "Dark mode" in menu_labels(app.view_menu)                  # dark mode lives under View
    # menu, undo/redo and formatting share one row
    assert app.undo_btn.winfo_parent() == app.fmtbar.winfo_parent()


def test_save_button_shows_unsaved_state(app, tmp_path):
    path = workbooks.plain(tmp_path / "p.xlsx")
    open_in(app, path)
    btn = app.save_btn
    assert btn.winfo_parent() == app.undo_btn.winfo_parent()        # on the toolbar row, by ☰
    assert app.save_icons["on"] is not app.save_icons["off"]
    assert btn.instate(["disabled"])                                 # nothing to save: grey
    app.set_cells([(0, 0, "changed")])
    assert not btn.instate(["disabled"])                             # unsaved: coloured, clickable
    btn.invoke()
    app.update()
    assert btn.instate(["disabled"]) and openpyxl.load_workbook(path)["Data"]["A1"].value == "changed"
    app.undo()
    assert not btn.instate(["disabled"])                             # file now differs again
    app.redo()
    assert btn.instate(["disabled"])                                 # back to what's on disk


def test_disk_icon_colours(app):
    import app_menu
    on = app_menu.disk_icon(app, 1.0, "#1a73e8")
    off = app_menu.disk_icon(app, 1.0, "#b5b5b5")
    assert on.get(2, 8) == (26, 115, 232) and off.get(2, 8) == (181, 181, 181)   # body colour
    assert on.get(6, 3) == (255, 255, 255)                                       # white shutter


def test_starts_with_blank_workbook(app):
    assert app.title() == "Simple Sheets - Untitled"
    assert list(app.sheets) == ["Sheet1"] and not app.dirty()
    app.set_cells([(0, 0, "hi")])
    assert app.dirty() and app.title() == "Simple Sheets - *Untitled"


def test_new_window_starts_new_process(app, monkeypatch):
    import simple_sheets
    calls = []
    monkeypatch.setattr(simple_sheets.subprocess, "Popen", lambda cmd, **kw: calls.append(cmd))
    app.event_generate("<Control-n>", when="now")
    app.new_window()
    assert calls and calls[-1][-1].endswith("simple_sheets.py")


def test_save_untitled_asks_for_name(app, tmp_path, monkeypatch):
    import simple_sheets
    dst = tmp_path / "new.xlsx"
    monkeypatch.setattr(simple_sheets.filedialog, "asksaveasfilename", lambda **k: str(dst))
    app.set_cells([(0, 0, "hello"), (0, 1, "=2*3")])
    assert app.save() and not app.dirty()
    ws = openpyxl.load_workbook(dst).active
    assert ws["A1"].value == "hello" and ws["B1"].value == "=2*3" and ws.title == "Sheet1"
    assert app.title() == "Simple Sheets - new.xlsx"


# ---------- open recent ----------

def test_recent_list(app, tmp_path):
    files = [workbooks.plain(tmp_path / f"f{i}.xlsx") for i in range(10)]
    for f in files:
        recent.add(f)
    recent.add(files[3])                      # re-opening moves it to the top
    items = recent.load()
    assert len(items) == recent.MAX and items[0] == os.path.abspath(files[3])
    assert items.count(os.path.abspath(files[3])) == 1


def test_open_records_and_menu_lists_recent(app, tmp_path):
    a = workbooks.plain(tmp_path / "alpha.xlsx")
    b = workbooks.sales(tmp_path / "beta.xlsx")
    open_in(app, a)
    open_in(app, b)
    import app_menu
    app_menu.fill_recent(app, app.recent_menu)
    labels = menu_labels(app.recent_menu)
    assert labels[0].startswith("1  beta.xlsx") and labels[1].startswith("2  alpha.xlsx")
    assert labels[-1] == "Clear list"
    app.recent_menu.invoke(1)                 # "2  alpha.xlsx"
    app.update()
    assert os.path.basename(app.path) == "alpha.xlsx" and list(app.sheets) == ["Data", "Other"]


def test_missing_recent_file_is_dropped(app, tmp_path):
    gone = tmp_path / "gone.xlsx"
    recent.add(gone)
    app.open_recent(str(gone))
    assert str(gone) not in recent.load() and app.path is None


def test_save_as_adds_to_recent(app, tmp_path, monkeypatch):
    import simple_sheets
    dst = tmp_path / "saved.xlsx"
    monkeypatch.setattr(simple_sheets.filedialog, "asksaveasfilename", lambda **k: str(dst))
    app.set_cells([(0, 0, "x")])
    app.save_as()
    assert recent.load()[0] == os.path.abspath(dst)


# ---------- sheets ----------

def test_add_sheet(app, tmp_path):
    open_in(app, workbooks.plain(tmp_path / "p.xlsx"))
    app.add_sheet()
    assert list(app.sheets) == ["Data", "Sheet3", "Other"]     # after the current sheet, unique name
    assert app.sheet == "Sheet3" and app.tabs.names() == list(app.sheets) and app.dirty()
    app.set_cells([(0, 0, "new data")])
    app.undo()
    app.undo()
    assert list(app.sheets) == ["Data", "Other"] and not app.dirty()
    app.redo()
    assert list(app.sheets) == ["Data", "Sheet3", "Other"] and app.sheet == "Sheet3"


def test_plus_button_adds_sheet(app):
    plus = next(w for w in app.tabs.winfo_children() if str(w.cget("text") if "text" in w.keys() else "") == "+")
    plus.invoke()
    app.update()
    assert list(app.sheets) == ["Sheet1", "Sheet2"] and app.tabs.names() == ["Sheet1", "Sheet2"]


@pytest.mark.parametrize("bad", ["", "a" * 32, "x/y", "[x]", "'x", "data"])
def test_rename_rejects_bad_names(app, tmp_path, bad):
    open_in(app, workbooks.plain(tmp_path / "p.xlsx"))
    assert not app.rename_sheet("Other", bad)
    assert list(app.sheets) == ["Data", "Other"]


def test_rename_updates_formulas_and_undo(app, tmp_path):
    g = open_in(app, workbooks.layout(tmp_path / "m.xlsx"))
    assert app.rename_sheet("Main", "Q1 Sales")
    assert list(app.sheets) == ["Q1 Sales", "Other"] and app.tabs.names() == ["Q1 Sales", "Other"]
    other = app.sheets["Other"]
    assert other.rows[0][0] == "='Q1 Sales'!D5" and other.rows[1][0] == "=SUM('Q1 Sales'!B2:B4)"
    app.show("Other")
    assert g.value(0, 0) == "140"
    app.undo()
    assert list(app.sheets) == ["Main", "Other"] and other.rows[0][0] == "=Main!D5" and not app.dirty()


def tab_xy(app, name):
    x1, x2 = app.tabs.tab_box(name)
    return (x1 + x2) // 2, app.tabs.nb.winfo_height() // 2


def test_inline_rename_on_double_click(app, tmp_path):
    open_in(app, workbooks.plain(tmp_path / "p.xlsx"))
    nb = app.tabs.nb
    x, y = tab_xy(app, "Other")
    assert nb.bind("<Double-1>")                              # Tk can't synthesize a double-click,
    app.tabs.rename_at(SimpleNamespace(x=x, y=y))             # so call its handler directly
    app.update()
    e = app.tabs.editor
    assert e is not None and e.winfo_ismapped() and e.get() == "Other"
    assert e.winfo_parent() == str(nb)                       # sits on the tab strip, not a pop-up
    assert abs(e.winfo_x() - app.tabs.tab_box("Other")[0]) <= 2
    e.delete(0, "end")
    e.insert(0, "Totals")
    key(app, "<Return>", widget=e)
    assert list(app.sheets) == ["Data", "Totals"] and app.tabs.editor is None
    assert app.tabs.names() == ["Data", "Totals"]


def test_inline_rename_escape_and_invalid(app, tmp_path):
    open_in(app, workbooks.plain(tmp_path / "p.xlsx"))
    app.tabs.start_rename("Other")
    app.update()
    e = app.tabs.editor
    e.delete(0, "end")
    e.insert(0, "Nope")
    key(app, "<Escape>", widget=e)
    assert list(app.sheets) == ["Data", "Other"] and app.tabs.editor is None
    app.tabs.start_rename("Other")
    app.update()
    e = app.tabs.editor
    e.delete(0, "end")
    e.insert(0, "data")                                       # duplicate (case-insensitive)
    key(app, "<Return>", widget=e)
    assert list(app.sheets) == ["Data", "Other"]
    assert any(d[0] == "showwarning" for d in app.dialogs)


def test_move_sheet_and_undo(app, tmp_path):
    path = workbooks.layout(tmp_path / "m.xlsx")
    open_in(app, path)
    app.add_sheet()                                           # Main, Sheet3, Other
    app.move_sheet("Other", 0)
    assert list(app.sheets) == ["Other", "Main", "Sheet3"] and app.tabs.names() == list(app.sheets)
    assert app.sheet == "Other" and app.dirty()
    app.undo()
    assert list(app.sheets) == ["Main", "Sheet3", "Other"]
    app.redo()
    assert app.save(), app.dialogs
    assert openpyxl.load_workbook(path).sheetnames == ["Other", "Main", "Sheet3"]


def test_drag_tab_to_reorder(app, tmp_path):
    open_in(app, workbooks.plain(tmp_path / "p.xlsx"))
    app.add_sheet()
    app.add_sheet()                                           # Data, Sheet4, Sheet3, Other
    app.update()
    order = list(app.sheets)
    nb = app.tabs.nb
    x0, y = tab_xy(app, order[0])
    end = app.tabs.tab_box(order[-1])[1] + 40                 # past the last tab
    nb.event_generate("<ButtonPress-1>", x=x0, y=y, when="now")
    for x in range(x0, end + 1, 10):
        nb.event_generate("<B1-Motion>", x=x, y=y, when="now")
        app.update()
    nb.event_generate("<ButtonRelease-1>", x=end, y=y, when="now")
    app.update()
    assert list(app.sheets) == order[1:] + order[:1] and app.tabs.names() == list(app.sheets)
    # and drag it back to the front
    x0, y = tab_xy(app, order[0])
    nb.event_generate("<ButtonPress-1>", x=x0, y=y, when="now")
    for x in range(x0, 0, -10):
        nb.event_generate("<B1-Motion>", x=x, y=y, when="now")
        app.update()
    nb.event_generate("<ButtonRelease-1>", x=2, y=y, when="now")
    app.update()
    assert list(app.sheets) == order
    app.undo()
    app.undo()
    assert list(app.sheets) == order


def test_click_without_drag_does_not_reorder(app, tmp_path):
    open_in(app, workbooks.plain(tmp_path / "p.xlsx"))
    nb = app.tabs.nb
    x, y = tab_xy(app, "Other")
    nb.event_generate("<ButtonPress-1>", x=x, y=y, when="now")
    nb.event_generate("<B1-Motion>", x=x + 2, y=y, when="now")     # tiny jitter, not a drag
    nb.event_generate("<ButtonRelease-1>", x=x + 2, y=y, when="now")
    app.update()
    assert list(app.sheets) == ["Data", "Other"] and app.hpos == 0


def test_delete_sheet(app, tmp_path):
    g = open_in(app, workbooks.layout(tmp_path / "m.xlsx"))
    assert app.delete_sheet("Main")
    assert list(app.sheets) == ["Other"] and app.sheet == "Other"
    assert app.sheets["Other"].rows[0][0] == "=#REF!" and g.value(0, 0) == "#REF!"
    assert not app.delete_sheet("Other")                   # the last sheet stays
    app.undo()
    assert list(app.sheets) == ["Main", "Other"] and app.sheets["Other"].rows[0][0] == "=Main!D5"


def test_edits_survive_rename(app, tmp_path):
    path = workbooks.plain(tmp_path / "p.xlsx")
    g = open_in(app, path)
    app.set_cells([(1, 1, "7")])
    app.rename_sheet("Data", "Prices")
    app.show("Prices")
    app.undo()        # undo rename
    app.undo()        # undo edit
    assert g.text(1, 1) == "3" and list(app.sheets) == ["Data", "Other"] and not app.dirty()
    app.redo()
    app.redo()
    assert app.save()
    wb = openpyxl.load_workbook(path)
    assert wb.sheetnames == ["Prices", "Other"] and wb["Prices"]["B2"].value == 7


def test_save_with_added_renamed_deleted_sheets(app, tmp_path):
    path = workbooks.layout(tmp_path / "m.xlsx")
    open_in(app, path)
    app.show("Main")
    app.add_sheet()                                  # Sheet3 after Main
    app.set_cells([(0, 0, "fresh"), (1, 0, "=SUM(Main!B2:B4)")])
    app.rename_sheet("Other", "Main")                # taken -> rejected
    app.rename_sheet("Main", "Other2")
    app.delete_sheet("Other")
    assert app.save(), app.dialogs
    assert not app.dirty()
    wb = openpyxl.load_workbook(path)
    assert wb.sheetnames == ["Other2", "Sheet3"]
    assert wb["Sheet3"]["A1"].value == "fresh" and wb["Sheet3"]["A2"].value == "=SUM(Other2!B2:B4)"
    assert wb["Other2"]["B3"].border.top.style == "thin" and "A7:B7" in str(wb["Other2"].merged_cells)
    app.undo()                                       # undo delete after save -> sheet comes back
    assert app.dirty()
    assert app.save(), app.dialogs
    wb = openpyxl.load_workbook(path)
    assert wb.sheetnames == ["Other2", "Sheet3", "Other"]
    assert wb["Other"]["A1"].value == "=Other2!D5"          # its content came back too
    assert wb["Other2"]["B3"].border.top.style == "thin"     # other sheets untouched


def test_csv_save_with_several_sheets_warns(app, tmp_path, monkeypatch):
    """Saving several sheets to CSV keeps only the current one and says so."""
    import simple_sheets
    dst = tmp_path / "out.csv"
    monkeypatch.setattr(simple_sheets.filedialog, "asksaveasfilename", lambda **k: str(dst))
    app.set_cells([(0, 0, "a")])
    app.add_sheet()
    app.set_cells([(0, 0, "b")])
    assert app.save_as()
    assert any(d[0] == "askokcancel" and d[1] == "Only one sheet" for d in app.dialogs)
    assert dst.read_text(encoding="utf-8").strip() == "b" and list(app.sheets) == ["out.csv"]
