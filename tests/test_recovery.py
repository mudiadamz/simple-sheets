"""Crash recovery: copies of unsaved work, cleaned up on save/close, offered back after a crash."""
import json
import os

import recovery
import workbooks
from conftest import open_in


def test_autosave_writes_and_cleans_up(app, tmp_path):
    path = workbooks.plain(tmp_path / "p.xlsx")
    open_in(app, path)
    xlsx, meta = recovery.paths(app.recovery_id)
    assert not app.autosave() and not os.path.exists(xlsx)          # nothing unsaved: no copy
    app.set_cells([(0, 0, "unsaved work")])
    assert app.autosave() and os.path.exists(xlsx)
    assert json.load(open(meta))["source"] == str(path)
    assert not app.autosave()                                         # unchanged since: not rewritten
    app.set_cells([(0, 1, "more")])
    assert app.autosave()
    assert app.save()
    assert not os.path.exists(xlsx) and not os.path.exists(meta)     # saved: copy removed


def test_undo_back_to_clean_removes_copy(app):
    app.set_cells([(0, 0, "x")])
    assert app.autosave()
    app.undo()
    app.autosave()
    assert not os.path.exists(recovery.paths(app.recovery_id)[0])


def fake_crash(app, monkeypatch):
    """Leave a copy behind as if its window had crashed (its process no longer runs)."""
    app.set_cells([(0, 0, "precious"), (1, 0, "=1+1")])
    app.autosave()
    xlsx, meta = recovery.paths(app.recovery_id)
    data = json.load(open(meta))
    data["pid"] = 999_999_999                       # no such process
    json.dump(data, open(meta, "w"))
    return xlsx


def test_offer_recovery_restores(app, tmp_path, monkeypatch):
    path = workbooks.plain(tmp_path / "orig.xlsx")
    open_in(app, path)
    xlsx = fake_crash(app, monkeypatch)
    app.recovery_id = recovery.new_id()             # this window is now "another" one
    assert [x for x, _ in recovery.orphans()] == [xlsx]
    app.new_document()
    app.offer_recovery()                            # stub says OK
    g = app.table
    assert g.text(0, 0) == "precious" and g.value(1, 0) == "2"
    assert app.path is None and app.dirty() and app.recovered_from == str(path)
    assert not os.path.exists(xlsx) and not recovery.orphans()
    assert os.path.exists(recovery.paths(app.recovery_id)[0])        # own copy straight away
    from simplecore import recent
    assert xlsx not in recent.load()


def test_offer_recovery_cancel_deletes(app, monkeypatch):
    xlsx = fake_crash(app, monkeypatch)
    app.recovery_id = recovery.new_id()
    app.new_document()
    import simple_sheets
    monkeypatch.setattr(simple_sheets.messagebox, "askokcancel", lambda *a, **k: False)
    app.offer_recovery()
    assert not os.path.exists(xlsx) and app.table.text(0, 0) == ""


def test_running_window_copy_not_offered(app):
    app.set_cells([(0, 0, "x")])
    app.autosave()
    assert recovery.orphans() == []                 # our own process is alive
