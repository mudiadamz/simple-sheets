"""Small differences between Windows, Linux and macOS: opening / printing files, the mouse wheel on Linux
(buttons 4 / 5), and ⌘ shortcuts on macOS (every Ctrl+key binding also works with Command)."""
import os
import re
import subprocess
import sys
import tkinter


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


def old_mac_wheel(platform=None, tk_version=None):
    """Tk 8.6 on macOS sends small wheel deltas (1, 2, 3…), not multiples of 120 like Tk 9 and Windows."""
    import tkinter
    return (platform or sys.platform) == "darwin" and (tk_version or tkinter.TkVersion) < 8.7


def touchpad_deltas(d):
    """<TouchpadScroll> (Tk 9, macOS two-finger scroll) packs the pixel deltas into %D: x high, y low 16 bits."""
    low = d & 0xFFFF
    return d >> 16, low - 0x10000 if low >= 0x8000 else low


def bind_touchpad(widget, handler):
    """Two-finger scroll on a Mac touchpad comes as <TouchpadScroll> in Tk 9 (not <MouseWheel>).
    False on an older Tk that doesn't know the event."""
    try:
        widget.bind("<TouchpadScroll>", handler)
        return True
    except Exception:    # TclError: bad event type
        return False


_MODS = re.compile(r"\b((?:Ctrl|Alt|Shift|Cmd)\+)+")


def mac_keys(text, platform=None):
    """A shortcut as a Mac shows it: "Ctrl+Shift+S" -> "⇧⌘S", "Alt+Down" -> "⌥Down" (other systems: as it is)."""
    if (platform or sys.platform) != "darwin" or not text or not isinstance(text, str):
        return text

    def symbols(m):
        mods = m.group(0)
        return "".join(sym for name, sym in (("Alt+", "⌥"), ("Shift+", "⇧"), ("Ctrl+", "⌘"), ("Cmd+", "⌘"))
                       if name in mods)
    return _MODS.sub(symbols, text)


def command_twin(sequence):
    """<Control-…> key / click binding -> its <Command-…> twin (None when there isn't one)."""
    if not isinstance(sequence, str) or "Control-" not in sequence or "MouseWheel" in sequence:
        return None
    return sequence.replace("Control-", "Command-")


def install_mac(platform=None):
    """macOS: ⌘ does whatever Ctrl does (every <Control-…> binding made from now on gets a <Command-…> twin),
    and menus show their shortcuts as ⌘ / ⇧ / ⌥. False elsewhere."""
    if (platform or sys.platform) != "darwin" or getattr(tkinter.Misc, "_mac_twins", False):
        return False
    misc, menu = tkinter.Misc, tkinter.Menu
    bind, bind_all, bind_class = misc.bind, misc.bind_all, misc.bind_class

    def twin(do, args, sequence, func, add):
        result = do(*args, sequence, func, add)
        other = command_twin(sequence)
        if func and other:
            try:
                do(*args, other, func, add)
            except tkinter.TclError:                   # a Tk without the Command modifier
                pass
        return result

    misc.bind = lambda self, sequence=None, func=None, add=None: twin(bind, (self,), sequence, func, add)
    misc.bind_all = lambda self, sequence=None, func=None, add=None: twin(bind_all, (self,), sequence, func, add)
    misc.bind_class = lambda self, className, sequence=None, func=None, add=None: twin(
        bind_class, (self, className), sequence, func, add)

    def keys(cnf, kw):
        cnf = dict(cnf or {}, **kw)
        if "accelerator" in cnf:
            cnf["accelerator"] = mac_keys(cnf["accelerator"], "darwin")
        return cnf
    add, insert, entryconfigure = menu.add, menu.insert, menu.entryconfigure
    menu.add = lambda self, itemType, cnf={}, **kw: add(self, itemType, keys(cnf, kw))
    menu.insert = lambda self, index, itemType, cnf={}, **kw: insert(self, index, itemType, keys(cnf, kw))
    menu.entryconfigure = menu.entryconfig = lambda self, index, cnf=None, **kw: (
        entryconfigure(self, index, keys(cnf, kw) if (cnf or kw) else None))
    misc._mac_twins = True
    return True


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
