"""Export to PDF / print."""
import re

import pdf_export
import workbooks
from conftest import open_in


def pdf_text(path):
    data = open(path, "rb").read()
    return data, re.findall(rb"\(((?:[^()\\]|\\.)*)\) Tj", data)


def test_export_values_formats_pages(app, tmp_path):
    g = open_in(app, workbooks.styled(tmp_path / "s.xlsx"))
    app.set_cells([(0, 6, "=1+2")])
    dst = str(tmp_path / "out.pdf")
    assert app.export_pdf(dst) == dst
    data, texts = pdf_text(dst)
    assert data.startswith(b"%PDF-1.4") and data.rstrip().endswith(b"%%EOF")
    assert b"3" in texts                                            # formula result, not the formula
    assert b"/Helvetica-Bold" in data
    assert data.count(b"/Type /Page ") == 1
    assert re.search(rb"/Count 1\b", data)
    assert g.sheet.styles and b" rg " in data                       # some fill / colour drawn
    xref = int(data.rsplit(b"startxref", 1)[1].split()[0])
    assert data[xref:xref + 4] == b"xref"                           # offsets right


def test_long_and_wide_sheets_split_into_pages(app, tmp_path):
    app.set_cells([(r, 0, f"row {r + 1}") for r in range(120)] + [(0, c, f"col{c}") for c in range(1, 30)])
    dst = str(tmp_path / "big.pdf")
    app.export_pdf(dst)
    data, texts = pdf_text(dst)
    pages = data.count(b"/Type /Page ")
    assert pages >= 4                                               # several pages down and across
    assert b"/MediaBox [0 0 842 595]" in data                       # wide: landscape
    assert b"row 120" in texts and b"col29" in texts


def test_hidden_rows_and_text_escaping(app, tmp_path):
    app.set_cells([(0, 0, "keep (1) \\ x"), (1, 0, "hidden row"), (2, 0, "café"), (2, 3, "end")])   # A1 spills right
    app.table.goto(1, 0)
    app.hide("row")
    dst = str(tmp_path / "h.pdf")
    app.export_pdf(dst)
    data, texts = pdf_text(dst)
    assert rb"keep \(1\) \\ x" in texts and b"hidden row" not in texts
    assert "café".encode("cp1252") in texts


def test_numbers_too_wide_show_hashes():
    m = lambda t, size, b, i: len(t) * size * 0.55   # noqa: E731
    assert pdf_export.clip("123456789012345", 30, 11, False, False, m, True).startswith("#")
    assert pdf_export.clip("long text here", 30, 11, False, False, m, False) == "long"


def test_print_uses_pdf(app, tmp_path, monkeypatch):
    sent = []
    monkeypatch.setattr(pdf_export.os, "startfile", lambda p, *a: sent.append((p, a)), raising=False)
    app.set_cells([(0, 0, "x")])
    assert app.print_sheet() == "printed"
    assert sent and sent[0][1] == ("print",) and sent[0][0].endswith(".pdf")


def test_menu_has_pdf_and_print(app):
    m = app.app_menu
    labels = [m.entrycget(i, "label") for i in range(m.index("end") + 1) if m.type(i) != "separator"]
    assert "Export to PDF…" in labels and "Print…" in labels
