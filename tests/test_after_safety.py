"""A delayed call (after / after_idle) made on a small window that closes before it runs crashes Tkinter
("'NoneType' object has no attribute 'remove'" in deletecommand). Such calls go on the main window."""
import glob
import os
import re
import sys
import time

import pytest

import screen_capture as sc

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


@pytest.fixture
def errors(app, monkeypatch):
    seen = []
    monkeypatch.setattr(app, "report_callback_exception", lambda *exc: seen.append(exc))
    return seen


def settle(app, ms=400):
    end = time.monotonic() + ms / 1000
    while time.monotonic() < end:
        app.update()
        time.sleep(0.01)


@pytest.mark.skipif(sys.platform != "win32", reason="Windows screen")
def test_capture_closes_its_layer_without_error(app, errors):
    got = []
    snip = sc.Snip(app, got.append, grab_fn=lambda x, y, w, h: sc.png(1, 1, b"\0\0\0"))
    app.update()
    snip.press(type("E", (), {"x": 5, "y": 5})())
    snip.release(type("E", (), {"x": 60, "y": 40})())
    settle(app)
    assert got and got[0] and not snip.win.winfo_exists() and errors == []


def test_tooltip_flashing_in_dark_mode(app, errors):
    from tkinter import ttk

    from format_bar import Tooltip
    app.set_theme(True)
    try:
        b = ttk.Button(app, text="x")
        b.pack()
        app.update()
        t = Tooltip(b, "hello")
        for _ in range(5):                      # shown and hidden before the title-bar change runs
            t.show()
            t.hide()
        settle(app)
        b.destroy()
    finally:
        app.set_theme(False)
    assert errors == []


def test_cell_editor_closed_while_autocomplete_waits(app, errors):
    g = app.table
    g.goto(0, 0)
    g.edit("=SU")
    app.update()
    g.editor.event_generate("<FocusOut>")      # autocomplete schedules a check on the editor...
    g.commit()                                  # ...which is closed right away
    settle(app)
    assert errors == []


def test_no_delayed_calls_on_short_lived_widgets():
    risky = re.compile(r"\b(win|button|widget|outer|tip|w|e|dialog|top|snip|editor)\.(after|after_idle)\(")
    found = []
    for p in glob.glob(os.path.join(ROOT, "*.py")):
        for n, line in enumerate(open(p, encoding="utf-8"), start=1):
            if risky.search(line) and "_root()" not in line:
                found.append(f"{os.path.basename(p)}:{n}: {line.strip()}")
    assert not found, "use widget._root().after(...) instead:\n" + "\n".join(found)
