"""Files in sync folders (OneDrive, Dropbox, Google Drive): service name, conflict copies."""
import json
import os

import cloud
import workbooks
from conftest import open_in


def test_roots_and_service(tmp_path):
    one, home = tmp_path / "OneDrive - Work", tmp_path / "home"
    (home / "Dropbox").mkdir(parents=True)
    one.mkdir()
    local = tmp_path / "local"
    (local / "Dropbox").mkdir(parents=True)
    other = tmp_path / "Box2"
    other.mkdir()
    (local / "Dropbox" / "info.json").write_text(json.dumps({"business": {"path": str(other)}}))
    env = {"OneDriveCommercial": str(one), "USERPROFILE": str(home), "LOCALAPPDATA": str(local)}
    assert cloud.service(str(one / "a" / "b.xlsx"), env) == "OneDrive"
    assert cloud.service(str(home / "Dropbox" / "b.xlsx"), env) == "Dropbox"
    assert cloud.service(str(other / "b.xlsx"), env) == "Dropbox"
    assert cloud.service(str(tmp_path / "OneDrive - Workish" / "b.xlsx"), env) is None   # not a prefix match
    assert cloud.service(str(tmp_path / "x.xlsx"), env) is None


def test_conflict_copies(tmp_path):
    for name in ("Budget.xlsx", "Budget-LAPTOP7.xlsx", "Budget (Ann's conflicted copy 2026-01-02).xlsx",
                 "Budget (1).xlsx", "Budget 2024.xlsx", "Budget-2024.xlsx", "Budgets.xlsx", "Budget-LAPTOP7.csv",
                 "budget-x.xlsx"):
        (tmp_path / name).write_text("x")
    found = [os.path.basename(p) for p in cloud.conflicts(str(tmp_path / "Budget.xlsx"))]
    assert found == ["Budget (1).xlsx", "Budget (Ann's conflicted copy 2026-01-02).xlsx", "Budget-LAPTOP7.xlsx"]


def test_check_warns_and_opens_newest(app, tmp_path, monkeypatch):
    folder = tmp_path / "OneDrive"
    folder.mkdir()
    path = workbooks.plain(folder / "Book.xlsx")
    copy = workbooks.plain(folder / "Book-DESKTOP1.xlsx")
    env = {"OneDrive": str(folder), "USERPROFILE": str(tmp_path / "nohome"), "LOCALAPPDATA": str(tmp_path)}
    opened = []
    import book_tabs
    monkeypatch.setattr(book_tabs, "open_in_tab", lambda a, p=None: opened.append(p))
    monkeypatch.setattr(cloud.messagebox if hasattr(cloud, "messagebox") else __import__("tkinter").messagebox,
                        "askyesno", lambda *a, **k: True)
    assert cloud.check(app, str(path), env) == [str(copy)]
    assert opened == [str(copy)] and "Synced by OneDrive" in app.status["text"]


def test_open_in_sync_folder_runs_check(app, tmp_path, monkeypatch):
    folder = tmp_path / "OneDrive"
    folder.mkdir()
    path = workbooks.plain(folder / "Book.xlsx")
    monkeypatch.setenv("OneDrive", str(folder))
    open_in(app, path)
    app.update()
    assert "Synced by OneDrive" in app.status["text"]
