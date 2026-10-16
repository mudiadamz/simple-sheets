"""PDF of all sheets / the selection, print preview, page breaks (view, manual, saved)."""
import re

import openpyxl

import pdf_export
import print_preview as pp
import workbooks
from conftest import open_in


def pages_in(path):
    data = open(path, "rb").read()
    return len(re.findall(rb"/Type /Page\b", data)), data


def long_sheet(app, n=120):
    app.set_cells([(r, 0, f"row {r + 1}") for r in range(n)] + [(r, 1, str(r)) for r in range(n)])


def test_pdf_all_sheets_and_selection(app, tmp_path):
    long_sheet(app)
    one = pdf_export.export(app, str(tmp_path / "a.pdf"))
    assert one >= 2                                                  # 120 rows: more than one A4 page
    app.table.select(0, 0, 9, 1)
    n = pdf_export.export(app, str(tmp_path / "s.pdf"), area=app.table.sel_cells())
    count, data = pages_in(tmp_path / "s.pdf")
    assert n == count == 1 and b"row 10" in data and b"row 11" not in data


def test_pdf_several_sheets(app, tmp_path):
    app.add_sheet()
    app.add_sheet()
    names = [n for n, s in app.sheets.items() if not s.hidden]
    for k, name in enumerate(names):
        app.show(name)
        app.set_cells([(0, 0, f"sheet {k}")])
    dst = app.export_pdf(str(tmp_path / "all.pdf"), "All sheets")
    count, data = pages_in(dst)
    assert count == len(names) and all(f"sheet {k}".encode() in data for k in range(len(names)))


def test_page_sizes_per_sheet(app, tmp_path):
    long_sheet(app, 3)
    pdf_export.write_pdf(str(tmp_path / "x.pdf"), [(b"", 595, 842), (b"", 842, 595)])
    data = open(tmp_path / "x.pdf", "rb").read()
    assert b"/MediaBox [0 0 595 842]" in data and b"/MediaBox [0 0 842 595]" in data


def test_manual_breaks_and_view(app, tmp_path):
    path = workbooks.plain(tmp_path / "b.xlsx")
    g = open_in(app, path)
    long_sheet(app, 30)
    assert pdf_export.plan(app)["row_starts"] == []                  # 30 rows fit on one page
    g.goto(10, 0)
    pp.insert_break(app)
    assert app.current.props["breaks"] == {"rows": [10], "cols": []}
    p = pdf_export.plan(app)
    assert [band[0][0] for band in p["pages"]] == [0, 10]
    assert pp.toggle(app) is True and app.page_breaks_var.get()
    g.redraw()
    assert g.cv.find_withtag("pagebreak")
    app.undo()
    assert not app.current.props.get("breaks")
    g.redraw()
    assert not g.cv.find_withtag("pagebreak")
    app.redo()
    assert app.save()
    ws = openpyxl.load_workbook(path).active
    assert [b.id for b in ws.row_breaks.brk] == [10]
    open_in(app, path)
    assert app.current.props["breaks"] == {"rows": [10], "cols": []}
    pp.remove_breaks(app)
    assert app.save()
    assert not openpyxl.load_workbook(path).active.row_breaks.brk


def test_automatic_breaks_drawn(app):
    long_sheet(app, 200)
    rows, cols, _ = pp.breaks(app)
    assert rows and rows[0] > 20
    pp.toggle(app, True)
    app.table.goto(rows[0], 0)
    app.table.redraw()
    assert app.table.cv.find_withtag("pagebreak")


def test_preview_matches_pdf_layout(app):
    """Wrapped text shows as several lines and long text is cut at the next filled cell — same as the PDF."""
    long = "Functional regression test for the sub domain management after the database migration"
    app.set_cells([(0, 0, "ID"), (0, 1, long), (0, 2, "next"), (1, 0, long), (1, 1, "B2")])
    s = app.current
    s.widths[1] = 20
    s.styles[(0, 1)] = {"wrap": True}
    w = pp.open_preview(app)
    w.update()
    w.draw()
    texts = [(w.cv.itemcget(i, "text"), w.cv.bbox(i)) for i in w.cv.find_all() if w.cv.type(i) == "text"]
    p, _ = w.pages[0]
    cells = pdf_export.page_cells(*p["pages"][0], p["widths"], p["heights"], p["merge_at"])
    box = next(b for r, c, b in cells if (r, c) == (0, 1))
    want = pdf_export.cell_text(0, 1, box, p["pages"][0][1], p["widths"], p["value"], p["style"], p["merge_at"],
                                p["measure"])
    assert len(want["lines"]) >= 3
    assert [t for t, _ in texts if t in long and t != long][:len(want["lines"])] == [x[2] for x in want["lines"]]
    a2 = next(t for t, _ in texts if long.startswith(t) and t != long and t not in [x[2] for x in want["lines"]])
    assert len(a2) < len(long)                         # A2 is cut where B2 starts
    a2_box = next(b for t, b in texts if t == a2)
    b2_box = next(b for t, b in texts if t == "B2")
    assert a2_box[2] <= b2_box[0] + 2
    w.destroy()


def test_preview_window(app, tmp_path, monkeypatch):
    long_sheet(app)
    w = pp.open_preview(app)
    w.update()
    total = len(w.pages)
    assert total >= 2 and w.label["text"] == f"Page 1 of {total}"
    texts = [w.cv.itemcget(i, "text") for i in w.cv.find_all() if w.cv.type(i) == "text"]
    assert "row 1" in texts
    w.go(1)
    assert w.label["text"] == f"Page 2 of {total}"
    w.go(99)
    assert w.i == total - 1
    app.table.select(0, 0, 4, 1)
    w.what.set("Selection")
    w.load()
    assert len(w.pages) == 1
    monkeypatch.setattr(pp.filedialog, "asksaveasfilename", lambda **k: str(tmp_path / "p.pdf"))
    assert w.export() == str(tmp_path / "p.pdf") and pages_in(tmp_path / "p.pdf")[0] == 1
    w.destroy()


def wide_sheet(app, cols=14, rows=60):
    app.set_cells([(r, c, f"value {r}-{c}") for r in range(rows) for c in range(cols)])


def test_ctrl_p_opens_the_dialog_first(app, monkeypatch):
    from conftest import key
    printed = []
    monkeypatch.setattr(pdf_export, "print_file", lambda path: printed.append(path) or "printed")
    app.set_cells([(0, 0, "x")])
    key(app, "<Control-p>")
    app.update()
    dialogs = [w for w in app.winfo_children() if w.winfo_class() == "Toplevel" and w.title() == "Print"]
    assert dialogs and printed == []                                  # nothing printed yet: the dialog shows
    w = dialogs[0]
    assert w.print_() == "printed" and len(printed) == 1               # Print… in the dialog prints
    w.destroy()


def test_orientation_and_scaling(app, tmp_path):
    wide_sheet(app)
    auto = pdf_export.plan(app)
    assert auto["page_w"] > auto["page_h"]                             # wide sheet: landscape by itself
    portrait = pdf_export.plan(app, orientation="portrait")
    assert portrait["page_w"] < portrait["page_h"]
    land = pdf_export.plan(app, orientation="landscape", fit="actual")
    assert land["scale"] == 1.0 and len(land["pages"]) > 1
    one = pdf_export.plan(app, orientation="portrait", fit="page")
    assert len(one["pages"]) == 1 and one["scale"] < 1                 # the whole sheet on one page
    width = pdf_export.plan(app, orientation="portrait", fit="width")
    assert len({tuple(p[1]) for p in width["pages"]}) == 1            # one page wide (all columns)
    pdf_export.export(app, str(tmp_path / "p.pdf"), orientation="portrait")
    assert b"/MediaBox [0 0 595 842]" in open(tmp_path / "p.pdf", "rb").read()


def test_dialog_options_used_and_remembered(app, tmp_path, monkeypatch):
    import settings
    wide_sheet(app, cols=3, rows=10)
    w = pp.open_preview(app)
    w.update()
    w.orientation.set("Landscape")
    w.fit.set("Fit the sheet on one page")
    w.options_changed()
    assert all(p["page_w"] > p["page_h"] for p, _ in w.pages) and len(w.pages) == 1
    assert settings.get("print_orientation") == "landscape" and settings.get("print_fit") == "page"
    monkeypatch.setattr(pp.filedialog, "asksaveasfilename", lambda **k: str(tmp_path / "o.pdf"))
    assert w.export() == str(tmp_path / "o.pdf")
    assert b"/MediaBox [0 0 842 595]" in open(tmp_path / "o.pdf", "rb").read()      # landscape PDF
    w.destroy()
    again = pp.open_preview(app)                                       # next time: same choices
    assert again.orientation.get() == "Landscape" and again.fit.get() == "Fit the sheet on one page"
    again.destroy()


def test_menu_print_and_export_open_the_dialog(app):
    menu = app.app_menu
    labels = {menu.entrycget(i, "label"): i for i in range(menu.index("end") + 1) if menu.type(i) == "command"}
    for label in ("Print…", "Export to PDF…", "Export all sheets to PDF…"):
        menu.invoke(labels[label])
        app.update()
        tops = [w for w in app.winfo_children() if w.winfo_class() == "Toplevel" and w.title() == "Print"]
        assert tops, label
        if label == "Export all sheets to PDF…":
            assert tops[-1].what.get() == "All sheets"
        for t in tops:
            t.destroy()
