"""Save that can't write its file (gone, its folder gone, no permission…) becomes Save as… a new file - not an
error box and a dead end: the workbook is still all here."""
import shutil

import openpyxl

from conftest import open_in


def _book(path):
    wb = openpyxl.Workbook()
    wb.active["A1"] = "kept"
    wb.save(path)
    return path


def _ask(monkeypatch, dst):
    import simple_sheets
    asked = []
    monkeypatch.setattr(simple_sheets.filedialog, "asksaveasfilename",
                        lambda **k: asked.append(k) or (str(dst) if dst else ""))
    return asked


def test_a_file_gone_is_saved_as_a_new_one(app, tmp_path, monkeypatch):
    path = _book(tmp_path / "gone.xlsx")
    open_in(app, path)
    app.set_cells([(1, 0, "new")])
    path.unlink()                                             # (deleted while it was open)
    dst = tmp_path / "again.xlsx"
    asked = _ask(monkeypatch, dst)
    assert app.save()
    assert len(asked) == 1 and "gone.xlsx" in asked[0]["title"] and "no longer there" in asked[0]["title"]
    assert asked[0]["initialfile"] == "gone.xlsx" and asked[0]["initialdir"] == str(tmp_path)
    assert not any("Save failed" in str(d) for d in app.dialogs)
    ws = openpyxl.load_workbook(dst).active
    assert ws["A1"].value == "kept" and ws["A2"].value == "new"   # (the whole workbook: nothing to patch into)
    assert app.path == str(dst)


def test_its_folder_gone_asks_where_with_no_dead_folder(app, tmp_path, monkeypatch):
    folder = tmp_path / "downloads"
    folder.mkdir()
    path = _book(folder / "b.xlsx")
    open_in(app, path)
    app.set_cells([(0, 1, "x")])
    shutil.rmtree(folder)
    dst = tmp_path / "b.xlsx"
    asked = _ask(monkeypatch, dst)
    assert app.save()
    assert "initialdir" not in asked[0]                         # (the dialog's own default, not a gone folder)
    assert openpyxl.load_workbook(dst).active["B1"].value == "x"


def test_a_write_error_becomes_save_as(app, tmp_path, monkeypatch):
    import sheet_io
    path = _book(tmp_path / "locked.xlsx")
    open_in(app, path)
    app.set_cells([(0, 0, "changed")])
    real = sheet_io.patch_xlsx

    def refuse(src, target, *a, **k):                         # (no permission any more on the original)
        if str(target) == str(path):
            raise PermissionError(13, "Permission denied", str(target))
        return real(src, target, *a, **k)
    monkeypatch.setattr(sheet_io, "patch_xlsx", refuse)
    dst = tmp_path / "copy.xlsx"
    asked = _ask(monkeypatch, dst)
    assert app.save()
    assert "permission" in asked[0]["title"]
    assert openpyxl.load_workbook(dst).active["A1"].value == "changed"


def test_cancelling_the_save_as_keeps_the_workbook_unsaved(app, tmp_path, monkeypatch):
    path = _book(tmp_path / "c.xlsx")
    open_in(app, path)
    app.set_cells([(0, 0, "edit")])
    path.unlink()
    _ask(monkeypatch, None)
    assert app.save() is False
    assert app.path == str(path) and app.dirty()


def test_save_as_errors_still_say_so(app, tmp_path, monkeypatch):
    """Save as… itself (a place you picked) failing: the error is shown - no loop of Save as dialogs."""
    import sheet_io
    monkeypatch.setattr(sheet_io, "write_xlsx", lambda *a, **k: (_ for _ in ()).throw(PermissionError(13, "Permission denied")))
    asked = _ask(monkeypatch, tmp_path / "n.xlsx")
    app.set_cells([(0, 0, "1")])
    assert app.save_as() is False
    assert len(asked) == 1 and any("Save failed" in str(d) for d in app.dialogs)
