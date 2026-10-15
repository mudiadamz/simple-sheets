""".ods files: formulas (OpenFormula), typed values, formatting, widths, heights, merges, hidden parts."""
import zipfile

import ods
import sheet_io as io
from conftest import open_in

CONTENT = """<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
 xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
 xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
 xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
 xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" office:version="1.2">
<office:automatic-styles>
 <number:number-style style:name="N2"><number:number number:decimal-places="2" number:grouping="true"/></number:number-style>
 <number:percentage-style style:name="N3"><number:number number:decimal-places="1"/><number:text>%</number:text></number:percentage-style>
 <number:date-style style:name="N4"><number:day number:style="long"/><number:text>/</number:text><number:month number:style="long"/><number:text>/</number:text><number:year number:style="long"/></number:date-style>
 <style:style style:name="co1" style:family="table-column"><style:table-column-properties style:column-width="1.5in"/></style:style>
 <style:style style:name="co2" style:family="table-column"><style:table-column-properties style:column-width="0.889in"/></style:style>
 <style:style style:name="ro1" style:family="table-row"><style:table-row-properties style:row-height="30pt" style:use-optimal-row-height="false"/></style:style>
 <style:style style:name="ta2" style:family="table"><style:table-properties table:display="false"/></style:style>
 <style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Heading">
  <style:table-cell-properties fo:background-color="#FFFF00" fo:wrap-option="wrap" style:vertical-align="middle"
   fo:border-bottom="1.5pt solid #0000FF" fo:border-top="none"/>
  <style:paragraph-properties fo:text-align="center"/>
  <style:text-properties fo:color="#FF0000" fo:font-size="14pt" style:font-name="Verdana"/>
 </style:style>
 <style:style style:name="ce2" style:family="table-cell" style:data-style-name="N2"/>
 <style:style style:name="ce3" style:family="table-cell" style:data-style-name="N3"/>
 <style:style style:name="ce4" style:family="table-cell" style:data-style-name="N4"/>
 <style:style style:name="ce5" style:family="table-cell"><style:table-cell-properties fo:border="0.74pt solid #000000"/></style:style>
</office:automatic-styles>
<office:body><office:spreadsheet>
 <table:table table:name="Data">
  <table:table-column table:style-name="co1"/>
  <table:table-column table:style-name="co2" table:number-columns-repeated="2"/>
  <table:table-column table:visibility="collapse"/>
  <table:table-row>
   <table:table-cell table:style-name="ce1" office:value-type="string" table:number-columns-spanned="2"><text:p>Big  <text:s text:c="2"/>title</text:p></table:table-cell>
   <table:covered-table-cell/>
   <table:table-cell office:value-type="boolean" office:boolean-value="true"><text:p>TRUE</text:p></table:table-cell>
  </table:table-row>
  <table:table-row table:style-name="ro1">
   <table:table-cell table:style-name="ce2" office:value-type="float" office:value="1234.5"><text:p>1,234.50</text:p></table:table-cell>
   <table:table-cell table:style-name="ce3" office:value-type="percentage" office:value="0.25"><text:p>25.0%</text:p></table:table-cell>
   <table:table-cell table:style-name="ce4" office:value-type="date" office:date-value="2024-03-05"><text:p>05/03/2024</text:p></table:table-cell>
  </table:table-row>
  <table:table-row table:visibility="collapse">
   <table:table-cell office:value-type="string"><text:p>hidden</text:p></table:table-cell>
  </table:table-row>
  <table:table-row>
   <table:table-cell table:formula="of:=SUM([.A2:.B2];1)" office:value-type="float" office:value="1235.75"><text:p>1235.75</text:p></table:table-cell>
   <table:table-cell table:formula="of:=IF([.A2]&gt;1;&quot;a;b&quot;;&quot;c&quot;)" office:value-type="string"><text:p>a;b</text:p></table:table-cell>
   <table:table-cell table:formula="of:=[$'Other sheet'.$A$1]*2" office:value-type="float" office:value="14"><text:p>14</text:p></table:table-cell>
   <table:table-cell table:style-name="ce5"/>
  </table:table-row>
  <table:table-row table:number-rows-repeated="1048570"><table:table-cell table:number-columns-repeated="1024"/></table:table-row>
 </table:table>
 <table:table table:name="Other sheet" table:style-name="ta2">
  <table:table-row><table:table-cell office:value-type="float" office:value="7"><text:p>7</text:p></table:table-cell></table:table-row>
 </table:table>
</office:spreadsheet></office:body></office:document-content>"""


def ods_file(path):
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("mimetype", "application/vnd.oasis.opendocument.spreadsheet")
        z.writestr("content.xml", CONTENT)
    return path


def test_formula_conversion():
    assert ods.excel_formula("of:=SUM([.A1:.B2];2)") == "=SUM(A1:B2,2)"
    assert ods.excel_formula('of:=IF([.A1]>1;"x;y";"")') == '=IF(A1>1,"x;y","")'
    assert ods.excel_formula("of:=[$Sheet2.A1:.B3]") == "=Sheet2!A1:B3"
    assert ods.excel_formula("of:=[$'My Sheet'.$A$1]") == "='My Sheet'!$A$1"
    assert ods.excel_formula("of:=COM.MICROSOFT.TEXTJOIN(\",\";1;[.A1:.A3])") == '=TEXTJOIN(",",1,A1:A3)'


def test_load_values_formulas_and_layout(tmp_path):
    book = io.load(str(ods_file(tmp_path / "t.ods")))
    s, other = book["Data"], book["Other sheet"]
    assert s.rows[0] == ["Big    title", "", "TRUE"]
    assert s.rows[1] == ["1234.5", "0.25", "2024-03-05"]
    assert s.rows[3][:3] == ["=SUM(A2:B2,1)", '=IF(A2>1,"a;b","c")', "='Other sheet'!$A$1*2"]
    assert len(s.rows) == 4                                    # the huge repeated empty row isn't expanded
    head = s.styles[(0, 0)]
    assert head == {"bg": "#ffff00", "wrap": True, "valign": "center", "border": {"bottom": ["medium", "#0000ff"]},
                    "align": "center", "fg": "#ff0000", "size": 14.0, "font": "Verdana"}
    assert s.styles[(1, 0)]["numfmt"] == "#,##0.00" and s.styles[(1, 1)]["numfmt"] == "0.0%"
    assert s.styles[(1, 2)]["numfmt"] == "dd/mm/yyyy"
    assert s.styles[(3, 3)]["border"]["left"] == ["thin", "#000000"]
    assert s.merges == [(0, 0, 0, 1)]
    assert s.widths[0] == round((108 / 0.75 - 5) / 7, 2) and 3 in s.hidden_cols
    assert s.heights == {1: 30.0} and s.hidden_rows == {2}
    assert other.hidden and not s.hidden


def test_ods_in_app(app, tmp_path):
    g = open_in(app, ods_file(tmp_path / "t.ods"))
    assert g.value(3, 0) == "1235.75" and g.value(3, 1) == "a;b" and g.value(3, 2) == "14"
    assert g.value(1, 1) == "0.25"
    from formatting import render
    assert render(g.value(1, 0), g.style(1, 0)["numfmt"]) == "1,234.50"
    assert render(g.value(1, 1), g.style(1, 1)["numfmt"]) == "25.0%"
    assert render(g.value(1, 2), g.style(1, 2)["numfmt"]) == "05/03/2024"
