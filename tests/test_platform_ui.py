"""Linux / macOS differences: mouse wheel buttons, ⌘ shortcuts, opening / printing files."""
import sys
import tkinter as tk

import pytest

import platform_ui


def test_linux_wheel(app):
    cv = tk.Canvas(app)
    cv.place(x=0, y=0)
    app.update()
    got = []
    platform_ui.bind_wheel(cv, lambda e: got.append(("v", e.delta)), lambda e: got.append(("h", e.delta)),
                           lambda e: got.append(("z", e.delta)), platform="linux")
    cv.event_generate("<Button-4>", x=1, y=1)
    cv.event_generate("<Button-5>", x=1, y=1)
    cv.event_generate("<Shift-Button-5>", x=1, y=1)
    assert got == [("v", 120), ("v", -120), ("h", -120)]
    cv.destroy()
    other = tk.Canvas(app)
    platform_ui.bind_wheel(other, got.append, platform="win32")
    assert not other.bind("<Button-4>")
    other.destroy()


def test_grid_scrolls_with_linux_wheel(app):
    g = app.table
    platform_ui.bind_wheel(g.cv, g.on_wheel, platform="linux")
    g.cv.event_generate("<Button-5>", x=5, y=5)
    assert g.top == 3


@pytest.mark.skipif(sys.platform == "darwin", reason="real macOS handles Command itself")
def test_command_keys_mirrored(app):
    w = tk.Frame(app)
    w.bind("<Control-s>", lambda e: None)
    w.bind("<Control-MouseWheel>", lambda e: None)
    try:
        n = platform_ui.mirror_command_keys(w, platform="darwin")
    except tk.TclError:
        pytest.skip("this Tk has no Command modifier")
    assert n == 1 and w.bind("<Command-s>") and not w.bind("<Command-MouseWheel>")
    assert platform_ui.mirror_command_keys(w, platform="win32") == 0
    w.destroy()


def test_print_and_open_fallbacks(monkeypatch):
    calls = []
    monkeypatch.setattr(platform_ui.sys, "platform", "linux")

    def run(args, **k):
        calls.append(args)
        raise platform_ui.subprocess.CalledProcessError(1, args)
    monkeypatch.setattr(platform_ui.subprocess, "run", run)
    monkeypatch.setattr(platform_ui.subprocess, "Popen", lambda args: calls.append(args))
    assert platform_ui.print_path("/tmp/x.pdf") == "opened"
    assert calls == [["lp", "/tmp/x.pdf"], ["xdg-open", "/tmp/x.pdf"]]
    monkeypatch.setattr(platform_ui.sys, "platform", "darwin")
    assert platform_ui.open_path("/tmp/x.pdf") and calls[-1] == ["open", "/tmp/x.pdf"]
