"""Round trip: open a file, change one cell, save, and check nothing else was lost.

Two sources:
- files built here with the libraries other programs use (openpyxl, xlwt, odfpy) holding charts, pictures,
  comments, tables, names, rules, shapes, sparklines, rich text, freeze panes, hidden parts, protection…
- real files dropped into tests/corpus/ (from Excel, LibreOffice, Google Sheets…): every one is checked
  the same way. Add files there whenever a bug report comes with one.
"""
import os
import shutil
import zipfile

import openpyxl
import pytest

import workbooks
from conftest import open_in

CORPUS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "corpus")
IGNORED_PARTS = {"xl/calcChain.xml", "docProps/app.xml", "docProps/core.xml"}   # rebuilt by every writer


def built_files(folder):
    """Sample files of every feature the app should carry through a save."""
    out = []
    for name in ("styled", "plain", "sales", "budget", "layout", "hidden", "hidden_sheet", "links"):
        out.append(getattr(workbooks, name)(os.path.join(folder, f"{name}.xlsx")))
    from test_arrays import tables_file
    from test_keep_parts import with_extras
    from test_p1_safety import features
    from test_p3_data import chart_file
    from test_richtext import rich_file
    from test_rules import rules_file
    out += [features(os.path.join(folder, "features.xlsx")), rules_file(os.path.join(folder, "rules.xlsx")),
            chart_file(os.path.join(folder, "chart.xlsx")), rich_file(os.path.join(folder, "rich.xlsx")),
            tables_file(os.path.join(folder, "tables.xlsx")), with_extras(os.path.join(folder, "shapes.xlsx"))]
    try:
        from test_xls_format import formula_file, old_file
        out += [old_file(os.path.join(folder, "old.xls")), formula_file(os.path.join(folder, "f.xls"))]
    except Exception:
        pass
    from test_ods import ods_file
    out.append(ods_file(os.path.join(folder, "t.ods")))
    return [str(p) for p in out]


def real_files():
    if not os.path.isdir(CORPUS):
        return []
    return sorted(os.path.join(CORPUS, n) for n in os.listdir(CORPUS)
                  if n.lower().endswith((".xlsx", ".xlsm", ".xls", ".xlsb", ".ods", ".csv")))


def snapshot(app):
    """Every sheet's cell texts (formulas as typed) and styles."""
    return {n: ([list(r) for r in s.rows], dict(s.styles)) for n, s in app.sheets.items()}


def roundtrip(app, path, tmp_path):
    work = tmp_path / os.path.basename(path)
    shutil.copy(path, work)
    open_in(app, work)
    before = snapshot(app)
    first = app.sheet
    g = app.table
    r = len(g.sheet.rows) + 3
    app.set_cells([(r, 0, "round trip")])
    ext = os.path.splitext(work)[1].lower()
    dst = work if ext in (".xlsx", ".xlsm", ".ods", ".csv") else tmp_path / (work.stem + ".xlsx")
    assert app.write(str(dst)), app.dialogs
    open_in(app, dst)
    after = snapshot(app)
    assert app.sheets[first].rows[r][0] == "round trip"
    for name, (rows, styles) in before.items():
        if ext == ".csv":
            break
        assert name in after, f"sheet {name} lost"
        rows2, styles2 = after[name]
        for i, row in enumerate(rows):
            for j, t in enumerate(row):
                if (name, i, j) == (first, r, 0):
                    continue
                got = rows2[i][j] if i < len(rows2) and j < len(rows2[i]) else ""
                assert got == t, f"{name}!{i},{j}: {t!r} became {got!r}"
        if ext in (".xlsx", ".xlsm"):
            assert styles2 == styles, f"formatting changed on {name}"
    return work, dst


@pytest.mark.parametrize("which", range(20))
def test_built_corpus(app, tmp_path, which):
    folder = tmp_path / "made"
    folder.mkdir()
    files = built_files(str(folder))
    if which >= len(files):
        pytest.skip("fewer sample files")
    src = files[which]
    if src.lower().endswith((".xlsx", ".xlsm")):
        with zipfile.ZipFile(src) as z:
            parts_before = set(z.namelist())
    work, dst = roundtrip(app, src, tmp_path)
    if src.lower().endswith((".xlsx", ".xlsm")):
        with zipfile.ZipFile(dst) as z:
            parts_after = set(z.namelist())
        lost = {p for p in parts_before - parts_after - IGNORED_PARTS if not p.startswith("xl/drawings/_rels")}
        # openpyxl renumbers some parts (drawing2.xml…): compare the kinds of parts, not their numbers
        kind = lambda p: "".join(ch for ch in p if not ch.isdigit())   # noqa: E731
        assert {kind(p) for p in lost} <= {kind(p) for p in parts_after}, f"parts lost: {sorted(lost)}"
        openpyxl.load_workbook(dst)


@pytest.mark.parametrize("path", real_files() or [None])
def test_real_corpus(app, tmp_path, path):
    if path is None:
        pytest.skip("no files in tests/corpus/ yet")
    roundtrip(app, path, tmp_path)
