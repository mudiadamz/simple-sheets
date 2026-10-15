"""Insert function dialog and Evaluate formula."""
import formula
import function_help


def test_every_function_has_help():
    missing = [n for n in formula.FUNCS if n not in function_help.HELP]
    assert missing == []
    assert "Dynamic arrays" in function_help.categories() and "Lambda" in function_help.categories()


def test_insert_function(app):
    d = function_help.InsertFunction(app)
    d.search.set("pmt")
    app.update()
    assert d.names == ["PMT"]
    assert "rate, nper" in d.sig.cget("text") and "Example" in d.example.cget("text")
    d.cat.set("Lambda")
    d.search.set("")
    d.fill()
    assert set(d.names) >= {"LET", "LAMBDA", "MAP"}
    d.search.set("vlookup")
    d.cat.set("All")
    d.fill()
    app.table.goto(0, 0)
    assert d.insert() == "VLOOKUP"
    assert app.table.editor.get() == "=VLOOKUP("


def test_evaluate_steps(app):
    app.set_cells([(0, 0, "2"), (0, 1, "3"), (0, 2, "=SUM(A1:B1)*A1+LEN(\"abc\")")])
    app.table.goto(0, 2)
    d = function_help.EvaluateDialog(app)
    parts = {p: v for _, p, v in d.rows}
    assert parts["A1:B1"] == "{2,3}" and parts["SUM(A1:B1)"] == "5" and parts["SUM(A1:B1)*A1"] == "10"
    assert d.rows[-1][2] == "13" and d.rows[-1][0] == 0
    assert len(d.tree.get_children()) == len(d.rows)


def test_unparse_round_trip():
    for f in ['=SUM(A1:B2,3)*-C1%', "=Sheet2!A1&\"x\"", "=IF(A1>1,TRUE,FALSE)", "=LAMBDA(x,x*2)(3)"]:
        tree = formula.Parser(f[1:]).parse()
        assert formula.Parser(function_help.unparse(tree)).parse() == tree
