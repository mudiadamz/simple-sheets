"""Small differences between Windows, Linux and macOS: opening / printing files, the mouse wheel on Linux
(buttons 4 / 5), and ⌘ shortcuts on macOS (every Ctrl+key binding also works with Command)."""
import os
import subprocess
import sys

# macOS: ⌘ twins, ⌘ / ⇧ / ⌥ labels, Tk 9 touchpad scroll (shared with the other apps)
from simplecore.macos import bind_touchpad, command_twin, mac_keys, old_mac_wheel, touchpad_deltas  # noqa: F401
from simplecore.macos import install as install_mac  # noqa: F401


def open_path(path):
    """Open a file with its program. True if something was started."""
    try:
        if sys.platform == "win32":
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
        return True
    except (OSError, AttributeError):
        return False


def print_path(path):
    """Print a file on the default printer: 'printed', 'opened' (shown instead, print from there) or None."""
    try:
        if sys.platform == "win32":
            os.startfile(path, "print")
        else:
            subprocess.run(["lp", path], check=True, capture_output=True, timeout=30)
        return "printed"
    except (OSError, AttributeError, subprocess.SubprocessError):
        return "opened" if open_path(path) else None


class _Wheel:
    def __init__(self, event, delta):
        self.__dict__.update(getattr(event, "__dict__", {}))
        self.delta = delta


def bind_wheel(widget, handler, shift_handler=None, ctrl_handler=None, platform=None):
    """Linux sends the wheel as buttons 4 (up) / 5 (down): turn them into <MouseWheel>-like events."""
    if (platform or sys.platform).startswith("linux"):
        for button, delta in (("4", 120), ("5", -120)):
            widget.bind(f"<Button-{button}>", lambda e, d=delta: handler(_Wheel(e, d)))
            if shift_handler:
                widget.bind(f"<Shift-Button-{button}>", lambda e, d=delta: shift_handler(_Wheel(e, d)))
            if ctrl_handler:
                widget.bind(f"<Control-Button-{button}>", lambda e, d=delta: ctrl_handler(_Wheel(e, d)))


def mirror_command_keys(widget, platform=None):
    """macOS: give every <Control-…> key binding of widget a <Command-…> twin (⌘S, ⌘Z…)."""
    if (platform or sys.platform) != "darwin":
        return 0
    n = 0
    for seq in widget.bind():
        if "Control-" in seq and "Button" not in seq and "MouseWheel" not in seq:
            twin = seq.replace("Control-", "Command-")
            if not widget.bind(twin):
                widget.bind(twin, widget.bind(seq))
                n += 1
    return n
