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


def test_touchpad_deltas_unpacked():
    pack = lambda dx, dy: (dx << 16) | (dy & 0xFFFF)   # noqa: E731  how Tk 9 builds %D
    assert platform_ui.touchpad_deltas(pack(0, -7)) == (0, -7)
    assert platform_ui.touchpad_deltas(pack(-3, 12)) == (-3, 12)
    assert platform_ui.touchpad_deltas(pack(40, 0)) == (40, 0)


def test_old_mac_wheel_only_on_tk86():
    assert platform_ui.old_mac_wheel("darwin", 8.6)
    assert not platform_ui.old_mac_wheel("darwin", 9.0)
    assert not platform_ui.old_mac_wheel("win32", 8.6)


class Pad:
    def __init__(self, dx, dy):
        self.delta = (dx << 16) | (dy & 0xFFFF)


def test_grid_two_finger_scroll(app):
    g = app.table
    app.set_cells([(r, c, "x") for r in range(200) for c in range(3)])
    app.update()
    step = g.rowh
    g.on_touchpad(Pad(0, -(step - 1)))                  # less than a row: waits
    assert g.top == 0
    g.on_touchpad(Pad(0, -1))                           # small pixel moves add up to a whole row
    assert g.top == 1 and g.left == 0
    g.on_touchpad(Pad(0, -2 * step))
    assert g.top == 3
    g.on_touchpad(Pad(0, 3 * step))                     # fingers the other way: back up
    assert g.top == 0
    g.on_touchpad(Pad(-g.colw(0), 0))                   # sideways: one column
    assert g.left == 1 and g.top == 0
    g.on_touchpad(Pad(0, 0))
    assert g.left == 1


def test_touchpad_bound_when_tk_has_it(app):
    g = app.table
    probe = tk.Frame(app)
    has_event = platform_ui.bind_touchpad(probe, lambda e: None)   # Tk 9: yes; Tk 8.6 (Linux CI): no event
    probe.destroy()
    assert bool(g.cv.bind("<TouchpadScroll>")) == has_event


def test_mac_shortcut_labels():
    assert platform_ui.mac_keys("Save (Ctrl+S)", "darwin") == "Save (⌘S)"
    assert platform_ui.mac_keys("Ctrl+Shift+S", "darwin") == "⇧⌘S"
    assert platform_ui.mac_keys("Alt+Shift+→", "darwin") == "⌥⇧→"
    assert platform_ui.mac_keys("Ctrl+Alt+5", "darwin") == "⌥⌘5"
    assert platform_ui.mac_keys("F9", "darwin") == "F9"
    assert platform_ui.mac_keys("Ctrl+S", "win32") == "Ctrl+S"


def test_command_twin():
    assert platform_ui.command_twin("<Control-s>") == "<Command-s>"
    assert platform_ui.command_twin("<Control-Shift-A>") == "<Command-Shift-A>"
    assert platform_ui.command_twin("<Control-Button-1>") == "<Command-Button-1>"
    assert platform_ui.command_twin("<Control-MouseWheel>") is None
    assert platform_ui.command_twin("<Key>") is None
    assert platform_ui.install_mac("win32") is False


@pytest.mark.skipif(sys.platform != "darwin", reason="the Command modifier exists only in Tk on macOS")
def test_mac_command_does_what_ctrl_does(app):
    g = app.table
    assert g.cv.bind("<Command-c>") and g.cv.bind("<Command-z>") or app.bind("<Command-z>")
    w = tk.Frame(app)
    got = []
    w.bind("<Control-k>", lambda e: got.append("k"))  # bound after the start: gets its twin too
    assert w.bind("<Command-k>")
    m = tk.Menu(app, tearoff=False)
    m.add_command(label="Save", accelerator="Ctrl+S")
    m.insert_command(0, label="Save as…", accelerator="Ctrl+Shift+S")
    assert m.entrycget(1, "accelerator") == "⌘S" and m.entrycget(0, "accelerator") == "⇧⌘S"
    w.destroy()
    m.destroy()


@pytest.mark.skipif(sys.platform != "darwin", reason="macOS only")
def test_unbound_command_key_does_not_type(app):
    g = app.table
    g.goto(0, 0)

    class E:
        char, state = "l", 0x8                       # ⌘L: no shortcut
    g.on_key(E())
    assert g.editor is None and g.text(0, 0) == ""
