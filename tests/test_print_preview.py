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
