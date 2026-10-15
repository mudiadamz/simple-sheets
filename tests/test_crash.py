"""Unexpected errors show a dialog with details and go to errors.log; the app keeps working."""
import os
import time

import crash


def test_error_in_a_callback(app):
    def boom():
        raise ZeroDivisionError("boom in a test")
    app.after(1, boom)
    for _ in range(20):
        time.sleep(0.01)
        app.update()
    assert "ZeroDivisionError: boom in a test" in app.last_error
    assert os.path.exists(crash.log_path())
    with open(crash.log_path(), encoding="utf-8") as f:
        assert "boom in a test" in f.read()
    dialogs = [w for w in app.winfo_children() if isinstance(w, crash.ErrorDialog)]
    assert len(dialogs) == 1
    dialogs[0].copy(dialogs[0].text)
    assert "boom in a test" in app.clipboard_get()
    app.after(1, boom)                                  # a second error: no second dialog on top
    for _ in range(20):
        time.sleep(0.01)
        app.update()
    assert len([w for w in app.winfo_children() if isinstance(w, crash.ErrorDialog)]) == 1
    dialogs[0].destroy()
    app.set_cells([(0, 0, "still works")])
    assert app.table.text(0, 0) == "still works"


def test_log_is_kept_small(tmp_path, monkeypatch):
    monkeypatch.setattr(crash, "MAX_LOG", 100)
    for i in range(5):
        crash.write_log("x" * 80 + str(i))
    with open(crash.log_path(), encoding="utf-8") as f:
        text = f.read()
    assert len(text) < 400 and "4" in text
