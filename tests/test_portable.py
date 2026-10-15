"""Portable mode: settings next to the program; no automatic install for a portable copy."""
import sys

import recent


def test_portable_dir(tmp_path):
    exe = tmp_path / "SimpleSheets.exe"
    exe.write_bytes(b"MZ")
    assert recent.portable_dir(str(exe)) is None
    (tmp_path / "portable.txt").write_text("")
    assert recent.portable_dir(str(exe)) == str(tmp_path / "SimpleSheetsData")
    (tmp_path / "portable.txt").unlink()
    (tmp_path / "SimpleSheetsData").mkdir()
    assert recent.portable_dir(str(exe)) == str(tmp_path / "SimpleSheetsData")


def test_config_dir_follows_portable_exe(tmp_path, monkeypatch):
    (tmp_path / "SimpleSheetsData").mkdir()
    monkeypatch.delenv("SIMPLE_SHEETS_HOME", raising=False)
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "executable", str(tmp_path / "SimpleSheets.exe"))
    assert recent.config_dir() == str(tmp_path / "SimpleSheetsData")
    import settings
    settings.set("lang", "en")
    assert (tmp_path / "SimpleSheetsData" / "settings.json").exists()


def test_config_dir_without_appdata(tmp_path, monkeypatch):
    monkeypatch.delenv("SIMPLE_SHEETS_HOME", raising=False)
    monkeypatch.delenv("APPDATA", raising=False)
    monkeypatch.setenv("XDG_CONFIG_HOME", str(tmp_path / "cfg"))
    assert recent.config_dir() == str(tmp_path / "cfg" / "SimpleSheets")


def test_portable_copy_does_not_self_install(app, tmp_path, monkeypatch):
    import hashlib
    import json

    import updates
    setup = tmp_path / "Setup.exe"
    setup.write_bytes(b"MZ")
    feed = tmp_path / "f.json"
    feed.write_text(json.dumps({"version": "99", "download": setup.as_uri(),
                                "sha256": hashlib.sha256(b"MZ").hexdigest(), "url": ""}))
    (tmp_path / "SimpleSheetsData").mkdir()
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "executable", str(tmp_path / "SimpleSheets.exe"))
    monkeypatch.setenv("SIMPLE_SHEETS_HOME", str(tmp_path / "home"))
    assert updates.check(app, feed.as_uri()) is True
    assert "download page" in str(app.dialogs[-1]) and updates.pending() is None
