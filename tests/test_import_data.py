"""Data → Import text / CSV wizard and Import table from web."""
import import_data as imp

PAGE = """<html><head><meta charset="utf-8"></head><body>
<table><tr><td>layout</td></tr></table>
<table><caption>Prices</caption>
<tr><th>Item</th><th>Price</th><th>Note</th></tr>
<tr><td>Pens</td><td>1,234</td><td>a<br>b</td></tr>
<tr><td colspan="2">Total &amp; more</td><td>x</td></tr>
</table></body></html>"""


def test_guessing():
    assert imp.guess_encoding(b"\xef\xbb\xbfa,b") == "utf-8-sig"
    assert imp.guess_encoding("é".encode("cp1252")) == "cp1252"
    assert imp.guess_encoding("é".encode("utf-8")) == "utf-8"
    assert imp.guess_delimiter("a;b;c\n1;2;3\n") == ";"
    assert imp.to_date("31/12/2024", "DMY") == "2024-12-31"
    assert imp.to_date("12/31/24", "MDY") == "2024-12-31"
    assert imp.to_date("hello", "DMY") == "hello" and imp.to_date("13/13/2024", "MDY") == "13/13/2024"


def test_parse_types_and_skip():
    raw = "junk line\nid;when;name;drop\n007;31.12.2024;Ann;x\n".encode("cp1252")
    rows, text = imp.parse(raw, "cp1252", ";", skip=1, types={0: "Text", 1: "Date (D/M/Y)", 3: "Skip"})
    assert rows == [["id", "when", "name"], ["007", "2024-12-31", "Ann"]]
    assert text == [0]


def test_wizard_new_sheet(app, tmp_path):
    p = tmp_path / "people.csv"
    p.write_bytes("id;name\n007;Zoë\n".encode("cp1252"))
    w = imp.import_csv(app, str(p))
    assert w.enc.get() == "cp1252" and w.delim.get() == "Semicolon"
    w.col.current(0)
    w.type.set("Text")
    w.set_type()
    assert w.tree.item(w.tree.get_children()[1])["values"][1] == "Zoë"
    name = w.ok()
    assert name == "people" and app.sheet == "people"
    s = app.sheets["people"]
    assert s.rows[1] == ["007", "Zoë"]
    assert s.styles[(1, 0)]["numfmt"] == "@" and s.styles[(0, 1)]["bold"]
    app.undo()
    assert "people" not in app.sheets


def test_wizard_at_cursor(app, tmp_path):
    p = tmp_path / "t.txt"
    p.write_text("a|b\n1|2\n", encoding="utf-8")
    w = imp.import_csv(app, str(p))
    w.new_sheet.set(False)
    app.table.goto(2, 1)
    w.ok()
    assert app.table.text(2, 1) == "a" and app.table.text(3, 2) == "2"
    app.undo()
    assert app.table.text(2, 1) == ""


def test_parse_tables():
    tables = imp.parse_tables(PAGE)
    assert len(tables) == 2 and tables[1]["caption"] == "Prices"
    assert tables[1]["rows"] == [["Item", "Price", "Note"], ["Pens", "1,234", "a\nb"], ["Total & more", "", "x"]]


def test_web_dialog(app, tmp_path):
    page = tmp_path / "p.html"
    page.write_text(PAGE, encoding="utf-8")
    d = imp.import_web(app)
    d.url.insert(0, page.as_uri())
    assert len(d.load()) == 2                                  # fetched through urllib (file://)
    assert "Prices" in d.list.get(1)
    d.list.selection_clear(0, "end")
    d.list.selection_set(1)
    name = d.ok()
    assert app.sheet == name and app.sheets[name].rows[1][1] == "1234"


def test_web_bad_address(app):
    d = imp.import_web(app)
    d.url.insert(0, "file:///no/such/file.html")
    assert d.load() == [] and app.dialogs[-1][0] == "showwarning"
    d.destroy()
