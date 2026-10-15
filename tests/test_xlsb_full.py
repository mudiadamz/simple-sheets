""".xlsb: formulas (both ways), formats, sizes, merges; Save as .xlsb (cross-checked with pyxlsb)."""
import struct
import zipfile

import pytest

import sheet_io as io
import xlsb
import xlsb_formula as xf
import xlsb_write
from conftest import open_in

FORMULAS = ["A1+B2*3", "SUM(A1:B5)", "$A$1&\"x\"\"y\"", "IF(A1>=2,\"big\",\"small\")", "-A1%", "(A1+1)/2^3",
            "ROUND(AVERAGE(B1:B9),2)", "SUM(A:A)", "SUM(2:3)", "COUNTIF($B$2:B9,\">1\")", "TRUE", "#N/A",
            "A$1<>$B2", "VLOOKUP(A1,C1:D9,2,FALSE)", "IF(A1,,1)", "70000.5", "MAX(1,A1,-2)", "INDEX(A1:B2,1,1)"]


@pytest.mark.parametrize("text", FORMULAS)
def test_formula_round_trip(text):
    rgce = xf.encode("=" + text)
    assert rgce is not None
    assert xf.decode(rgce) == text


def test_formula_classes():
    # a range given to SUM is a reference; A1 in =A1+1 is a value
    assert xf.encode("=SUM(A1:B2)")[0] == 0x25 and xf.encode("=A1+1")[0] == 0x44
    assert xf.encode("=SUM(A1:B2*2)")[0] == 0x45


def test_unsupported_formulas_return_none():
    for text in ("=Sheet2!A1", "=XLOOKUP(1,A:A,B:B)", "=myName*2", "={1,2}", "=Table1[Col]"):
        assert xf.encode(text) is None
    assert xf.decode(b"\x23\x01\x00\x00\x00") is None              # PtgName: not read -> keep the value


def test_shared_and_3d_refs():
    # PtgRefN relative to the cell (row +1, col -1), PtgRef3d to the 2nd sheet
    refn = bytes([0x4C]) + struct.pack("<iH", 1, (-1 & 0x3FFF) | 0xC000)
    assert xf.decode(refn, base=(4, 3)) == "C6"
    ref3d = bytes([0x5A]) + struct.pack("<HIH", 1, 0, 0xC000)
    assert xf.decode(ref3d, sheets=["One", "Two words"]) == "'Two words'!A1"


def book():
    s = io.Sheet([["Name", "Qty", "Price", "Total"], ["Pens", "2", "1.5", "=B2*C2"], ["Ink", "3", "4", "=B3*C3"],
                  ["", "", "Sum", "=SUM(D2:D3)"], ["2024-03-05", "TRUE", "#N/A", "=IF(B2>1,\"many\",\"few\")"]])
    s.styles = {(0, 0): {"bold": True, "bg": "#ffff00", "fg": "#ff0000", "size": 14.0, "font": "Arial"},
                (1, 2): {"numfmt": "0.00", "align": "right"}, (2, 2): {"italic": True, "underline": True},
                (3, 3): {"border": {"top": ["thin", "#000000"], "bottom": ["double", "#0000ff"]}, "wrap": True,
                         "valign": "top"},
                (4, 0): {"numfmt": "yyyy-mm-dd"}, (6, 6): {"bg": "#00ff00"}}
    s.widths = {0: 20.0, 3: 12.5}
    s.heights = {2: 30.0}
    s.merges = [(5, 0, 5, 2)]
    s.hidden_rows, s.hidden_cols = {7}, {5}
    s.reset_structure()
    h = io.Sheet([["secret"]])
    h.hidden = True
    h.reset_structure()
    return {"Data": s, "Hidden": h}


def test_write_and_read_back(app, tmp_path):
    app.sheets.clear()
    app.sheets.update(book())
    app.engine.set_book(app.sheets)
    dst = tmp_path / "out.xlsb"
    assert xlsb_write.write(str(dst), app, app.sheets) == 0
    back = xlsb.load(str(dst))
    d = back["Data"]
    assert list(back) == ["Data", "Hidden"] and back["Hidden"].hidden and not d.hidden
    assert d.rows[1] == ["Pens", "2", "1.5", "=B2*C2"] and d.rows[3][3] == "=SUM(D2:D3)"
    assert d.rows[4][:3] == ["2024-03-05", "TRUE", "#N/A"] and d.rows[4][3] == '=IF(B2>1,"many","few")'
    assert d.cache[(1, 3)] == "3" and d.cache[(3, 3)] == "15" and d.cache[(4, 3)] == "many"
    assert d.styles[(0, 0)] == {"bold": True, "bg": "#ffff00", "fg": "#ff0000", "size": 14.0, "font": "Arial"}
    assert d.styles[(1, 2)] == {"numfmt": "0.00", "align": "right"}
    assert d.styles[(2, 2)] == {"italic": True, "underline": True}
    assert d.styles[(3, 3)] == {"border": {"top": ["thin", "#000000"], "bottom": ["double", "#0000ff"]},
                                "wrap": True, "valign": "top"}
    assert d.styles[(6, 6)] == {"bg": "#00ff00"}
    assert d.widths[0] == 20.0 and d.widths[3] == 12.5 and d.heights == {2: 30.0}
    assert d.merges == [(5, 0, 5, 2)] and d.hidden_rows == {7} and 5 in d.hidden_cols


def test_pyxlsb_reads_our_file(app, tmp_path):
    pyxlsb = pytest.importorskip("pyxlsb")
    app.sheets.clear()
    app.sheets.update(book())
    app.engine.set_book(app.sheets)
    dst = tmp_path / "out.xlsb"
    xlsb_write.write(str(dst), app, app.sheets)
    with pyxlsb.open_workbook(str(dst)) as wb:
        assert wb.sheets == ["Data", "Hidden"]
        with wb.get_sheet("Data") as sh:
            rows = [[c.v for c in row] for row in sh.rows()]
    assert rows[0][:4] == ["Name", "Qty", "Price", "Total"]
    assert rows[1][:4] == ["Pens", 2.0, 1.5, 3.0] and rows[3][3] == 15.0 and rows[4][3] == "many"


def test_zip_parts(app, tmp_path):
    app.sheets.clear()
    app.sheets.update(book())
    dst = tmp_path / "out.xlsb"
    xlsb_write.write(str(dst), app, app.sheets)
    with zipfile.ZipFile(dst) as z:
        assert z.testzip() is None
        names = set(z.namelist())
        ct = z.read("[Content_Types].xml").decode()
    assert {"xl/workbook.bin", "xl/styles.bin", "xl/sharedStrings.bin", "xl/worksheets/sheet2.bin"} <= names
    assert "application/vnd.ms-excel.worksheet" in ct


def test_save_as_xlsb_in_app(app, tmp_path):
    from test_p3_files import xlsb_file
    g = open_in(app, xlsb_file(tmp_path / "b.xlsb"))
    app.set_cells([(3, 0, "5"), (3, 1, "=A4*2"), (3, 2, "=Hidden!A1")])
    assert app.save()                                                  # .xlsb is saved in place now
    assert app.dialogs[-1][0] == "showwarning" and "1 formula" in str(app.dialogs[-1])
    back = xlsb.load(str(tmp_path / "b.xlsb"))
    assert back["Data"].rows[3] == ["5", "=A4*2", "h"]                 # the cross-sheet one kept as its value
    assert back["Data"].rows[0] == ["shared text", "12", "3.25", "TRUE"]
    g = open_in(app, tmp_path / "b.xlsb")
    assert g.value(3, 1) == "10"


def test_unsupported_parts_warned(app, tmp_path, monkeypatch):
    from test_p3_files import xlsb_file
    path = xlsb_file(tmp_path / "c.xlsb")
    with zipfile.ZipFile(path, "a") as z:
        z.writestr("xl/drawings/drawing1.xml", "<x/>")
    assert xlsb.unsupported(str(path)) == ["pictures or charts"]
    open_in(app, path)
    app.set_cells([(5, 0, "x")])
    import simple_sheets
    monkeypatch.setattr(simple_sheets.messagebox, "askokcancel", lambda *a, **k: False)
    assert app.save() is False
