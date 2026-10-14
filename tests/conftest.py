import os
import sys
import time

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

if sys.platform == "win32":  # same pixel math as the real app
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass


def make_app():
    """A new App window; Tk sometimes fails to start when many windows open in one process."""
    import simple_sheets
    for attempt in range(5):
        try:
            return simple_sheets.App()
        except simple_sheets.tk.TclError:
            if attempt == 4:
                raise
            time.sleep(0.1)


@pytest.fixture
def app(monkeypatch, tmp_path):
    """A real (invisible) app window. Dialogs are stubbed so tests never block;
    the recent-files list lives in the test's temp folder."""
    monkeypatch.setenv("SIMPLE_SHEETS_HOME", str(tmp_path / "home"))
    import simple_sheets
    mb = simple_sheets.messagebox
    dialogs = []   # every stubbed dialog, for assertions / debugging: (kind, title, message)

    def stub(kind, answer):
        return lambda *a, **k: (dialogs.append((kind,) + a), answer)[1]
    monkeypatch.setattr(mb, "askyesnocancel", stub("askyesnocancel", False))
    monkeypatch.setattr(mb, "askokcancel", stub("askokcancel", True))
    for name in ("showerror", "showwarning", "showinfo"):
        monkeypatch.setattr(mb, name, stub(name, None))
    a = make_app()
    a.dialogs = dialogs
    a.opened_state = a.state()   # checked by a test; then use a normal window (steadier focus)
    a.state("normal")
    # invisible AND off-screen: an invisible window still catches real mouse clicks
    a.geometry(f"{int(1000 * a.scale)}x{int(450 * a.scale)}-20000-20000")
    a.attributes("-alpha", 0.0)
    a.update()
    yield a
    a.destroy()


def open_in(app, path):
    app.open_file(str(path))
    app.update()
    return app.table


def key(app, sequence, widget=None, **kw):
    """Send a key/mouse event to a widget (default: the grid). Key events need real keyboard focus."""
    w = widget or app.table.cv
    for _ in range(50):   # a just-opened window may take a moment to get focus
        w.focus_force()
        app.update()
        if app.focus_get() is w:
            break
        time.sleep(0.02)
    w.event_generate(sequence, when="now", **kw)
    app.update()
