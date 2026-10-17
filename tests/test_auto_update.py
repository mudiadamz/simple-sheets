"""Automatic updates: feed parsing, verified download, install at next start; silent setup."""
import hashlib
import json
import sys

import pytest
from simplecore import updates


@pytest.fixture
def home(tmp_path, monkeypatch):
    monkeypatch.setenv("SIMPLE_SHEETS_HOME", str(tmp_path / "home"))
    return tmp_path


def make_feed(tmp_path, version="99.0.0", good=True, style="plain"):
    setup = tmp_path / "SimpleSheets-Setup.exe"
    setup.write_bytes(b"MZ new setup")
    digest = hashlib.sha256(setup.read_bytes()).hexdigest() if good else "0" * 64
    if style == "plain":
        data = {"version": version, "url": "https://example.invalid/r", "download": setup.as_uri(), "sha256": digest}
    else:                                   # a GitHub release: setup asset + checksum asset
        side = tmp_path / "SimpleSheets-Setup.exe.sha256"
        side.write_text(f"{digest}  SimpleSheets-Setup.exe\n")
        data = {"tag_name": f"v{version}", "html_url": "https://example.invalid/r", "body": "Notes",
                "assets": [{"name": "SimpleSheets-Setup.exe", "browser_download_url": setup.as_uri()},
                           {"name": "SimpleSheets-Setup.exe.sha256", "browser_download_url": side.as_uri()}]}
    feed = tmp_path / "latest.json"
    feed.write_text(json.dumps(data))
    return feed, digest


def test_info_plain_and_github(home):
    feed, digest = make_feed(home)
    i = updates.info(feed.as_uri())
    assert i["version"] == "99.0.0" and i["sha256"] == digest and i["download"].endswith("Setup.exe")
    feed, digest = make_feed(home, style="github")
    i = updates.info(feed.as_uri())
    assert i["version"] == "99.0.0" and i["sha256"] == digest
    feed.write_text(json.dumps({"tag_name": "v2", "body": f"SHA-256: {'a' * 64}", "assets": [
        {"name": "x-setup.exe", "browser_download_url": "file:///nowhere"}]}))
    assert updates.info(feed.as_uri())["sha256"] == "a" * 64


def test_download_verifies_hash(home):
    feed, _ = make_feed(home)
    path = updates.download(updates.info(feed.as_uri()))
    assert open(path, "rb").read() == b"MZ new setup" and updates.pending()["version"] == "99.0.0"
    updates.clear_pending()
    feed, _ = make_feed(home, good=False)
    with pytest.raises(ValueError):
        updates.download(updates.info(feed.as_uri()))
    assert updates.pending() is None
    with pytest.raises(ValueError):
        updates.download({"version": "9", "download": "file:///x", "sha256": ""})


def test_apply_pending_runs_setup_once(home):
    feed, _ = make_feed(home)
    path = updates.download(updates.info(feed.as_uri()))
    ran = []
    assert updates.apply_pending(run=ran.append) is True
    assert ran == [[path, "/S", "/relaunch"]]
    assert updates.apply_pending(run=ran.append) is False and len(ran) == 1      # once only


def test_apply_pending_rejects_tampered_or_old(home):
    feed, _ = make_feed(home)
    path = updates.download(updates.info(feed.as_uri()))
    open(path, "ab").write(b"evil")
    ran = []
    assert updates.apply_pending(run=ran.append) is False and ran == [] and updates.pending() is None
    feed, _ = make_feed(home, version="0.0.1")
    updates.download(updates.info(feed.as_uri()))
    assert updates.apply_pending(run=ran.append) is False and ran == []


def test_check_downloads_when_allowed(app, home, monkeypatch):
    feed, _ = make_feed(home)
    app.allow_dev_update = True
    try:
        assert updates.check(app, feed.as_uri()) is True
        assert app.dialogs[-1][0] == "askokcancel" and "next time" in str(app.dialogs[-1])
        assert updates.pending()["version"] == "99.0.0" and "downloaded" in app.status["text"]
        assert updates.check(app, feed.as_uri()) is True and app.dialogs[-1][0] == "showinfo"   # already ready
    finally:
        del app.allow_dev_update
        updates.clear_pending()


def test_background_check(app, home):
    feed, _ = make_feed(home)
    asked = []
    orig = updates.check
    updates.check = lambda a, url=None, quiet=False, known=None: asked.append(known["version"])
    try:
        updates.check_in_background(app, feed.as_uri())
        import time
        end = time.monotonic() + 5
        while not asked and time.monotonic() < end:
            app.update()
            time.sleep(0.05)
    finally:
        updates.check = orig
    assert asked == ["99.0.0"]


@pytest.mark.skipif(sys.platform != "win32", reason="Windows registry")
def test_silent_install(tmp_path, monkeypatch):
    from simplecore import installer
    fake = tmp_path / "payload.exe"
    fake.write_bytes(b"MZ app")
    monkeypatch.setattr(installer, "payload", lambda: str(fake))
    monkeypatch.setattr(installer, "_notify_explorer", lambda: None)
    base = r"Software\SimpleSheetsTest"
    opts = installer.parse_args(["/S", f"/D={tmp_path / 'app'}", "/noopenwith"])
    assert opts["silent"] and not opts["open_with"] and opts["dir"] == str(tmp_path / "app")
    log = tmp_path / "setup.log"
    try:
        assert installer.silent_install(opts, base=base, shortcuts=False, log=str(log)) == 0
        assert (tmp_path / "app" / installer.EXE_NAME).exists() and installer.existing(base)["dir"] == str(tmp_path / "app")
        assert "installed" in log.read_text()
        monkeypatch.setattr(installer, "in_use", lambda d: True)            # a copy that never closes
        assert installer.silent_install(opts, base=base, shortcuts=False, wait=0.6, log=str(log)) == 1
        assert "install failed" in log.read_text()
    finally:
        installer._delete_tree("", base)
