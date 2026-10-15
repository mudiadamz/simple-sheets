"""Shapes, text boxes and sparklines survive a save (openpyxl alone drops them)."""
import re
import zipfile

import openpyxl

import keep_parts
import workbooks
from conftest import open_in

DRAWING = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
           'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
           '<xdr:twoCellAnchor><xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row>'
           '<xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff>'
           '<xdr:row>8</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
           '<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="2" name="TextBox 1"/><xdr:cNvSpPr txBox="1"/>'
           '</xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>'
           '<xdr:txBody><a:bodyPr/><a:p><a:r><a:t>Hello note</a:t></a:r></a:p></xdr:txBody></xdr:sp>'
           '<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>')
SPARK = ('<extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" '
         'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">'
         '<x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">'
         '<x14:sparklineGroup><x14:sparklines><x14:sparkline><xm:f>Sheet!A1:B1</xm:f><xm:sqref>C1</xm:sqref>'
         '</x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext></extLst>')


def with_extras(path):
    """A workbook made by openpyxl, then given a text box and a sparkline group by hand (like Excel would)."""
    workbooks.plain(path)
    with zipfile.ZipFile(path) as z:
        files = {n: z.read(n) for n in z.namelist()}
    sheet = "xl/worksheets/sheet1.xml"
    xml = files[sheet].decode()
    if 'xmlns:r=' not in xml:
        xml = xml.replace("<worksheet ", '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/'
                                         'relationships" ', 1)
    xml = xml.replace("</worksheet>", '<drawing r:id="rIdShape"/>' + SPARK + "</worksheet>")
    files[sheet] = xml.encode()
    files["xl/worksheets/_rels/sheet1.xml.rels"] = (
        '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/'
        '2006/relationships"><Relationship Id="rIdShape" Type="http://schemas.openxmlformats.org/officeDocument/'
        '2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>').encode()
    files["xl/drawings/drawing1.xml"] = DRAWING.encode()
    ct = files["[Content_Types].xml"].decode().replace(
        "</Types>", '<Override PartName="/xl/drawings/drawing1.xml" '
                    'ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>')
    files["[Content_Types].xml"] = ct.encode()
    with zipfile.ZipFile(path, "w") as z:
        for n, b in files.items():
            z.writestr(n, b)
    return path


def parts(path):
    with zipfile.ZipFile(path) as z:
        return {n: z.read(n).decode("utf-8", "replace") for n in z.namelist()}


def test_shapes_and_sparklines_kept(app, tmp_path):
    path = with_extras(tmp_path / "s.xlsx")
    open_in(app, path)
    assert app.lost_parts == []                         # nothing to warn about any more
    app.set_cells([(0, 0, "edited")])
    assert app.save()
    p = parts(path)
    drawing = next(v for k, v in p.items() if k.startswith("xl/drawings/drawing"))
    assert "Hello note" in drawing and "<xdr:row>4</xdr:row>" in drawing
    assert "sparklineGroups" in p["xl/worksheets/sheet1.xml"] and "<drawing r:id=" in p["xl/worksheets/sheet1.xml"]
    wb = openpyxl.load_workbook(path)                  # still a valid workbook
    assert wb.active["A1"].value == "edited"


def test_shapes_move_with_inserted_rows(app, tmp_path):
    path = with_extras(tmp_path / "s.xlsx")
    open_in(app, path)
    app.restructure("row", 0, 2)                       # two rows above the text box
    assert app.save()
    drawing = next(v for k, v in parts(path).items() if k.startswith("xl/drawings/drawing"))
    assert "<xdr:row>6</xdr:row>" in drawing and "<xdr:row>10</xdr:row>" in drawing


def test_shape_on_deleted_rows_goes(tmp_path):
    src = with_extras(tmp_path / "s.xlsx")
    dst = tmp_path / "d.xlsx"
    openpyxl.load_workbook(src).save(dst)              # what openpyxl alone does: the box is gone
    assert not any("Hello note" in v for v in parts(dst).values())
    assert keep_parts.restore(str(src), str(dst), maps={"Data": (lambda r: None if 3 <= r <= 9 else r, None)}) \
        == 0
    assert not any("Hello note" in v for v in parts(dst).values())


def test_restore_into_existing_drawing(tmp_path):
    src = with_extras(tmp_path / "s.xlsx")
    dst = tmp_path / "d.xlsx"
    openpyxl.load_workbook(src).save(dst)
    assert keep_parts.restore(str(src), str(dst)) == 2   # the text box and the sparkline group
    p = parts(dst)
    assert any("Hello note" in v for k, v in p.items() if k.startswith("xl/drawings/"))
    assert re.search(r'drawing\+xml', p["[Content_Types].xml"])
